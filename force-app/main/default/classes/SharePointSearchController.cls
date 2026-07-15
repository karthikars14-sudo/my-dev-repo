public with sharing class SharePointSearchController {
    private static final String NAMED_CREDENTIAL = 'callout:SharePoint_Search';
    private static final String GRAPH_SEARCH_PATH = '/v1.0/search/query';
    private static final String SITE_PATH = 'https://home3925.sharepoint.com/sites/DemoSite';
    private static final Integer ROW_LIMIT = 50;

    // Required by Graph for searches using application permissions. Must match the geographic
    // region your Microsoft 365 tenant's data lives in - check Microsoft 365 admin center > Settings
    // > Org settings > Organization profile > Data location. Common values: NAM, EUR, APC, AUS.
    private static final String SEARCH_REGION = 'NAM';

    private static final String MP_AUTHOR = 'Author';
    // Maps each search field to the SharePoint managed property it queries.
    // If your Search Schema uses different managed property names, update these constants only.
    private static final String MP_DOCUMENT_TYPE = 'DocumentType';
    private static final String MP_PROJECT_ID = 'ProjectID';
    private static final String MP_OPPORTUNITY_ID = 'OpportunityID';
    private static final String DOWNLOAD_URL_FIELD = '@microsoft.graph.downloadUrl';

    public class SharePointDocumentResult {
        @AuraEnabled
        public String title;
        @AuraEnabled
        public String author;
        @AuraEnabled
        public String documentType;
        @AuraEnabled
        public String projectId;
        @AuraEnabled
        public String opportunityId;
        @AuraEnabled
        public String fileUrl;
        @AuraEnabled
        public String lastModified;
        @AuraEnabled
        public String driveId;
        @AuraEnabled
        public String itemId;
    }

    @AuraEnabled
    public static List<SharePointDocumentResult> searchDocuments(
        String documentName,
        String documentType,
        String projectId,
        String opportunityId,
        String author
    ) {
        String queryText = buildQueryText(
            documentName,
            documentType,
            projectId,
            opportunityId,
            author
        );

        HttpRequest req = new HttpRequest();
        req.setEndpoint(NAMED_CREDENTIAL + GRAPH_SEARCH_PATH);
        req.setMethod('POST');
        req.setHeader('Content-Type', 'application/json');
        req.setBody(JSON.serialize(buildRequestBody(queryText)));

        HttpResponse res;
        try {
            res = new Http().send(req);
        } catch (Exception e) {
            throw new AuraHandledException(
                'Unable to reach Microsoft Graph: ' + e.getMessage()
            );
        }

        if (res.getStatusCode() != 200) {
            throw new AuraHandledException(
                'SharePoint search failed (' +
                    res.getStatusCode() +
                    '): ' +
                    res.getStatus()
            );
        }

        List<SharePointDocumentResult> results = parseResults(res.getBody());
        enrichWithCustomFields(results);
        return results;
    }

    // Fetches the direct, no-auth-required download URL for one document. Called on demand (when
    // the user clicks Download for a specific row) rather than eagerly for every search result -
    // doing it eagerly for up to ROW_LIMIT results meant up to 50 extra callouts per search, which
    // was triggering rapid repeated OAuth token refreshes against Azure AD and intermittent
    // "Unable to fetch the OAuth token" failures. One callout per user click avoids that.
    @AuraEnabled
    public static String getDownloadUrl(String driveId, String itemId) {
        if (String.isBlank(driveId) || String.isBlank(itemId)) {
            throw new AuraHandledException('Missing document identifiers.');
        }

        HttpRequest req = new HttpRequest();
        req.setEndpoint(
            NAMED_CREDENTIAL +
                '/v1.0/drives/' +
                driveId +
                '/items/' +
                itemId +
                '?$select=' +
                EncodingUtil.urlEncode(DOWNLOAD_URL_FIELD, 'UTF-8')
        );
        req.setMethod('GET');

        HttpResponse res;
        try {
            res = new Http().send(req);
        } catch (Exception e) {
            throw new AuraHandledException(
                'Unable to reach Microsoft Graph: ' + e.getMessage()
            );
        }

        if (res.getStatusCode() != 200) {
            throw new AuraHandledException(
                'Unable to get a download link (' +
                    res.getStatusCode() +
                    '): ' +
                    res.getStatus()
            );
        }

        Map<String, Object> item = (Map<String, Object>) JSON.deserializeUntyped(
            res.getBody()
        );
        String downloadUrl = (String) item.get(DOWNLOAD_URL_FIELD);
        if (String.isBlank(downloadUrl)) {
            throw new AuraHandledException(
                'Microsoft Graph did not return a download link for this document.'
            );
        }
        return downloadUrl;
    }

    // NOTE: deliberately NOT requesting a "fields" array here. Confirmed twice against a live
    // tenant (including with the correct managed-property casing) that specifying one on a
    // driveItem search blanks out the entire result set. Microsoft's own documented example of
    // "fields" only shows it working for entityTypes=externalItem, never driveItem - this looks
    // like a real API limitation rather than a config mistake. Custom columns are fetched
    // separately afterward via enrichWithCustomFields().
    private static Map<String, Object> buildRequestBody(String queryText) {
        Map<String, Object> searchRequest = new Map<String, Object>{
            'entityTypes' => new List<String>{ 'driveItem' },
            'query' => new Map<String, Object>{ 'queryString' => queryText },
            'region' => SEARCH_REGION,
            'from' => 0,
            'size' => ROW_LIMIT
        };
        return new Map<String, Object>{
            'requests' => new List<Object>{ searchRequest }
        };
    }

    private static final String BATCH_PATH = '/v1.0/$batch';
    // Microsoft Graph's JSON batching endpoint accepts at most 20 sub-requests per call.
    private static final Integer BATCH_CHUNK_SIZE = 20;

    // Fetches DocumentType/ProjectID/OpportunityID for every result in one or a few $batch calls
    // (not one callout per result) via each document's listItem?$expand=fields - the standard,
    // well-documented way to read SharePoint list column values through Graph. Doing this eagerly
    // per result would repeat the earlier "many callouts in one transaction" problem that caused
    // OAuth token-refresh throttling; batching keeps it to at most ROW_LIMIT/20 extra callouts
    // (so at most 3 with the current ROW_LIMIT of 50), regardless of how many results come back.
    // Best-effort: any failure here leaves the affected results' custom fields blank rather than
    // failing the whole search.
    private static void enrichWithCustomFields(
        List<SharePointDocumentResult> results
    ) {
        Integer index = 0;
        while (index < results.size()) {
            Integer endExclusive = Math.min(
                index + BATCH_CHUNK_SIZE,
                results.size()
            );
            List<SharePointDocumentResult> chunk = new List<SharePointDocumentResult>();
            for (Integer i = index; i < endExclusive; i++) {
                chunk.add(results[i]);
            }
            fetchCustomFieldsForChunk(chunk);
            index = endExclusive;
        }
    }

    private static void fetchCustomFieldsForChunk(
        List<SharePointDocumentResult> chunk
    ) {
        List<Object> batchRequests = new List<Object>();
        for (Integer i = 0; i < chunk.size(); i++) {
            SharePointDocumentResult result = chunk[i];
            if (
                String.isBlank(result.driveId) || String.isBlank(result.itemId)
            ) {
                continue;
            }
            batchRequests.add(
                new Map<String, Object>{
                    'id' => String.valueOf(i),
                    'method' => 'GET',
                    'url' => '/drives/' +
                    result.driveId +
                    '/items/' +
                    result.itemId +
                    '/listItem?$expand=fields'
                }
            );
        }
        if (batchRequests.isEmpty()) {
            return;
        }

        HttpRequest req = new HttpRequest();
        req.setEndpoint(NAMED_CREDENTIAL + BATCH_PATH);
        req.setMethod('POST');
        req.setHeader('Content-Type', 'application/json');
        req.setBody(
            JSON.serialize(
                new Map<String, Object>{ 'requests' => batchRequests }
            )
        );

        HttpResponse res;
        try {
            res = new Http().send(req);
        } catch (Exception e) {
            return;
        }
        if (res.getStatusCode() != 200) {
            return;
        }

        Map<String, Object> body = (Map<String, Object>) JSON.deserializeUntyped(
            res.getBody()
        );
        List<Object> responses = (List<Object>) body.get('responses');
        if (responses == null) {
            return;
        }

        for (Object responseObj : responses) {
            Map<String, Object> response = (Map<String, Object>) responseObj;
            Object statusObj = response.get('status');
            if (statusObj == null || (Integer) statusObj != 200) {
                continue;
            }

            Integer chunkIndex = Integer.valueOf((String) response.get('id'));
            if (chunkIndex < 0 || chunkIndex >= chunk.size()) {
                continue;
            }

            Map<String, Object> responseBody = (Map<String, Object>) response.get(
                'body'
            );
            Map<String, Object> fields = responseBody == null
                ? null
                : (Map<String, Object>) responseBody.get('fields');
            if (fields == null) {
                continue;
            }

            SharePointDocumentResult result = chunk[chunkIndex];
            result.documentType = (String) fields.get(MP_DOCUMENT_TYPE);
            result.projectId = (String) fields.get(MP_PROJECT_ID);
            result.opportunityId = (String) fields.get(MP_OPPORTUNITY_ID);
        }
    }

    // Matches Document Name on the "filename" managed property (the document's actual file name, as
    // opposed to the separate, often-empty "Title" list column) with a trailing wildcard, so "test"
    // matches "testdocument.docx". Note: SharePoint/Graph KQL only supports prefix wildcards (word*),
    // not leading or mid-word wildcards - "test" won't match "mydocument-test.docx" since "test" isn't
    // a token prefix there. True arbitrary substring search isn't supported by the search index.
    //
    // Document Type/Project ID/Opportunity ID match exactly (no wildcard) against the custom managed
    // properties configured in your SharePoint Search Schema (see the MP_* constants above).
    private static String buildQueryText(
        String documentName,
        String documentType,
        String projectId,
        String opportunityId,
        String author
    ) {
        List<String> clauses = new List<String>();
        clauses.add('Path:"' + SITE_PATH + '*"');
        if (String.isNotBlank(documentName)) {
            clauses.add('filename:"' + escapeKql(documentName) + '*"');
        }
        if (String.isNotBlank(documentType)) {
            clauses.add(
                MP_DOCUMENT_TYPE + ':"' + escapeKql(documentType) + '"'
            );
        }
        if (String.isNotBlank(projectId)) {
            clauses.add(MP_PROJECT_ID + ':"' + escapeKql(projectId) + '"');
        }
        if (String.isNotBlank(opportunityId)) {
            clauses.add(
                MP_OPPORTUNITY_ID + ':"' + escapeKql(opportunityId) + '"'
            );
        }
        if (String.isNotBlank(author)) {
            clauses.add(MP_AUTHOR + ':"' + escapeKql(author) + '*"');
        }
        return String.join(clauses, ' AND ');
    }

    private static String escapeKql(String value) {
        return value.replace('"', '');
    }

    private static List<SharePointDocumentResult> parseResults(
        String jsonBody
    ) {
        List<SharePointDocumentResult> results = new List<SharePointDocumentResult>();

        Map<String, Object> root = (Map<String, Object>) JSON.deserializeUntyped(
            jsonBody
        );
        List<Object> searchResponses = (List<Object>) root.get('value');
        if (searchResponses == null) {
            return results;
        }

        for (Object searchResponseObj : searchResponses) {
            Map<String, Object> searchResponse = (Map<String, Object>) searchResponseObj;
            List<Object> hitsContainers = (List<Object>) searchResponse.get(
                'hitsContainers'
            );
            if (hitsContainers == null) {
                continue;
            }
            for (Object hitsContainerObj : hitsContainers) {
                Map<String, Object> hitsContainer = (Map<String, Object>) hitsContainerObj;
                List<Object> hits = (List<Object>) hitsContainer.get('hits');
                if (hits == null) {
                    continue;
                }
                for (Object hitObj : hits) {
                    SharePointDocumentResult result = toDocumentResult(
                        (Map<String, Object>) hitObj
                    );
                    if (result != null) {
                        results.add(result);
                    }
                }
            }
        }

        return results;
    }

    // webUrl opens SharePoint's browser viewer (requires the user to be signed in, and
    // previews/edits rather than downloading) - it's used as the initial/fallback link. The real
    // download link is fetched on demand via getDownloadUrl() using the driveId/itemId captured here.
    //
    // Returns null (skip) for folders and Site Pages. A driveItem only has a "folder" facet when it
    // IS a folder - this is the documented driveItem schema, more reliable than filtering on an
    // unconfirmed KQL managed property name. Modern SharePoint pages (Site Pages) are stored as
    // .aspx driveItems in the site's SitePages library, which is itself a document library, so they
    // come back from a driveItem search alongside real uploaded files - the .aspx extension is what
    // distinguishes a page from an actual document.
    private static SharePointDocumentResult toDocumentResult(
        Map<String, Object> hit
    ) {
        Map<String, Object> resource = (Map<String, Object>) hit.get(
            'resource'
        );

        if (resource == null || resource.get('folder') != null) {
            return null;
        }

        String name = (String) resource.get('name');
        if (name != null && name.toLowerCase().endsWith('.aspx')) {
            return null;
        }

        SharePointDocumentResult result = new SharePointDocumentResult();
        result.title = name;
        result.fileUrl = (String) resource.get('webUrl');
        result.lastModified = (String) resource.get('lastModifiedDateTime');
        result.itemId = (String) resource.get('id');

        Map<String, Object> parentReference = (Map<String, Object>) resource.get(
            'parentReference'
        );
        result.driveId = parentReference == null
            ? null
            : (String) parentReference.get('driveId');

        Map<String, Object> createdBy = (Map<String, Object>) resource.get(
            'createdBy'
        );
        Map<String, Object> createdByUser = createdBy == null
            ? null
            : (Map<String, Object>) createdBy.get('user');
        result.author = createdByUser == null
            ? null
            : (String) createdByUser.get('displayName');

        // documentType/projectId/opportunityId are filled in afterward by enrichWithCustomFields()
        // via a separate $batch call - the search response itself doesn't carry them (see
        // buildRequestBody()).

        return result;
    }
}
