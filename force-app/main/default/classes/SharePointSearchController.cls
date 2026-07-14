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
    private static final String DOWNLOAD_URL_FIELD = '@microsoft.graph.downloadUrl';

    public class SharePointDocumentResult {
        @AuraEnabled
        public String title;
        @AuraEnabled
        public String author;
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
        String author
    ) {
        String queryText = buildQueryText(documentName, author);

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

        return parseResults(res.getBody());
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

    // Matches on the "filename" managed property (the document's actual file name, as opposed to
    // the separate, often-empty "Title" list column) with a trailing wildcard, so "test" matches
    // "testdocument.docx". Note: SharePoint/Graph KQL only supports prefix wildcards (word*), not
    // leading or mid-word wildcards - "test" won't match "mydocument-test.docx" since "test" isn't
    // a token prefix there. True arbitrary substring search isn't supported by the search index.
    private static String buildQueryText(String documentName, String author) {
        List<String> clauses = new List<String>();
        clauses.add('Path:"' + SITE_PATH + '*"');
        if (String.isNotBlank(documentName)) {
            clauses.add('filename:"' + escapeKql(documentName) + '*"');
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

        return result;
    }
}
