@IsTest
private class SharePointSearchControllerTest {
    // Handles both callouts searchDocuments() can make: the POST to /search/query, and the
    // follow-up POST to /$batch that fetches custom field values per result.
    private class GraphApiMock implements HttpCalloutMock {
        private final String searchResponseBody;
        private final Map<String, Map<String, String>> customFieldsByItemId;
        private HttpRequest capturedSearchRequest;
        private HttpRequest capturedBatchRequest;

        GraphApiMock(
            String searchResponseBody,
            Map<String, Map<String, String>> customFieldsByItemId
        ) {
            this.searchResponseBody = searchResponseBody;
            this.customFieldsByItemId = customFieldsByItemId;
        }

        public HttpResponse respond(HttpRequest req) {
            HttpResponse res = new HttpResponse();
            if (req.getEndpoint().endsWith('/search/query')) {
                capturedSearchRequest = req;
                res.setStatusCode(200);
                res.setStatus('OK');
                res.setBody(searchResponseBody);
                return res;
            }

            // $batch call
            capturedBatchRequest = req;
            Map<String, Object> body = (Map<String, Object>) JSON.deserializeUntyped(
                req.getBody()
            );
            List<Object> subRequests = (List<Object>) body.get('requests');
            List<Object> subResponses = new List<Object>();
            for (Object subRequestObj : subRequests) {
                Map<String, Object> subRequest = (Map<String, Object>) subRequestObj;
                String id = (String) subRequest.get('id');
                String url = (String) subRequest.get('url');
                String itemId = url.substringAfter('/items/')
                    .substringBefore('/listItem');
                Map<String, String> customFields = customFieldsByItemId.get(
                    itemId
                );
                if (customFields == null) {
                    subResponses.add(
                        new Map<String, Object>{ 'id' => id, 'status' => 404 }
                    );
                } else {
                    subResponses.add(
                        new Map<String, Object>{
                            'id' => id,
                            'status' => 200,
                            'body' => new Map<String, Object>{
                                'fields' => customFields
                            }
                        }
                    );
                }
            }
            res.setStatusCode(200);
            res.setStatus('OK');
            res.setBody(
                JSON.serialize(
                    new Map<String, Object>{ 'responses' => subResponses }
                )
            );
            return res;
        }
    }

    private class DownloadUrlMock implements HttpCalloutMock {
        private final Integer statusCode;
        private final String body;
        private HttpRequest capturedRequest;

        DownloadUrlMock(Integer statusCode, String body) {
            this.statusCode = statusCode;
            this.body = body;
        }

        public HttpResponse respond(HttpRequest req) {
            capturedRequest = req;
            HttpResponse res = new HttpResponse();
            res.setStatusCode(statusCode);
            res.setStatus(statusCode == 200 ? 'OK' : 'Error');
            res.setBody(body);
            return res;
        }
    }

    private static String hitJson(
        String itemId,
        String driveId,
        String name,
        String authorDisplayName,
        String lastModified,
        String webUrl
    ) {
        return '{' +
            '"resource":{' +
            '"@odata.type":"#microsoft.graph.driveItem",' +
            '"id":"' +
            itemId +
            '",' +
            '"parentReference":{"driveId":"' +
            driveId +
            '"},' +
            '"name":"' +
            name +
            '",' +
            '"webUrl":"' +
            webUrl +
            '",' +
            '"lastModifiedDateTime":"' +
            lastModified +
            '",' +
            '"createdBy":{"user":{"displayName":"' +
            authorDisplayName +
            '"}}' +
            '}}';
    }

    private static String folderHitJson(
        String itemId,
        String driveId,
        String name,
        String webUrl
    ) {
        return '{' +
            '"resource":{' +
            '"@odata.type":"#microsoft.graph.driveItem",' +
            '"id":"' +
            itemId +
            '",' +
            '"parentReference":{"driveId":"' +
            driveId +
            '"},' +
            '"name":"' +
            name +
            '",' +
            '"webUrl":"' +
            webUrl +
            '",' +
            '"folder":{"childCount":3}' +
            '}}';
    }

    private static String sitePageHitJson(
        String itemId,
        String driveId,
        String name,
        String webUrl
    ) {
        return '{' +
            '"resource":{' +
            '"@odata.type":"#microsoft.graph.driveItem",' +
            '"id":"' +
            itemId +
            '",' +
            '"parentReference":{"driveId":"' +
            driveId +
            '"},' +
            '"name":"' +
            name +
            '",' +
            '"webUrl":"' +
            webUrl +
            '"' +
            '}}';
    }

    private static String responseWithHits(List<String> hits) {
        return '{"value":[{"searchTerms":[],"hitsContainers":[{"total":' +
            hits.size() +
            ',"moreResultsAvailable":false,"hits":[' +
            String.join(hits, ',') +
            ']}]}]}';
    }

    @IsTest
    static void searchDocumentsEnrichesResultsWithCustomFieldsViaBatch() {
        String hit = hitJson(
            'item-1',
            'drive-1',
            'testdocument.docx',
            'Jane Doe',
            '2026-06-01T10:00:00Z',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx'
        );
        Map<String, Map<String, String>> customFieldsByItemId = new Map<String, Map<String, String>>{
            'item-1' => new Map<String, String>{
                'DocumentType' => 'Contract',
                'ProjectID' => '12345',
                'OpportunityID' => '00654321'
            }
        };
        Test.setMock(
            HttpCalloutMock.class,
            new GraphApiMock(
                responseWithHits(new List<String>{ hit }),
                customFieldsByItemId
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'test',
            'Contract',
            '12345',
            '00654321',
            'Jane Doe'
        );
        Test.stopTest();

        System.assertEquals(
            1,
            results.size(),
            'Expected exactly one mapped document result'
        );
        System.assertEquals('testdocument.docx', results[0].title);
        System.assertEquals('Jane Doe', results[0].author);
        System.assertEquals('Contract', results[0].documentType);
        System.assertEquals('12345', results[0].projectId);
        System.assertEquals('00654321', results[0].opportunityId);
        System.assertEquals(
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx',
            results[0].fileUrl
        );
        System.assertEquals('drive-1', results[0].driveId);
        System.assertEquals('item-1', results[0].itemId);
    }

    @IsTest
    static void searchDocumentsLeavesCustomFieldsBlankWhenBatchLookupFails() {
        String hit = hitJson(
            'item-2',
            'drive-1',
            'nodata.docx',
            'Someone',
            '2026-06-01T10:00:00Z',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/nodata.docx'
        );
        Test.setMock(
            HttpCalloutMock.class,
            new GraphApiMock(
                responseWithHits(new List<String>{ hit }),
                new Map<String, Map<String, String>>()
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'nodata',
            null,
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(1, results.size());
        System.assertEquals(
            'nodata.docx',
            results[0].title,
            'Base fields should still populate'
        );
        System.assertEquals(
            null,
            results[0].documentType,
            'Custom fields stay blank when the batch lookup has no data for this item'
        );
    }

    @IsTest
    static void searchDocumentsExcludesFolders() {
        String folderHit = folderHitJson(
            'folder-1',
            'drive-1',
            'Archive',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/Archive'
        );
        String fileHit = hitJson(
            'item-4',
            'drive-1',
            'testdocument.docx',
            'Jane Doe',
            '2026-06-01T10:00:00Z',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx'
        );
        Test.setMock(
            HttpCalloutMock.class,
            new GraphApiMock(
                responseWithHits(new List<String>{ folderHit, fileHit }),
                new Map<String, Map<String, String>>()
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'test',
            null,
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(
            1,
            results.size(),
            'Expected the folder hit to be filtered out'
        );
        System.assertEquals('testdocument.docx', results[0].title);
    }

    @IsTest
    static void searchDocumentsExcludesSitePages() {
        String sitePageHit = sitePageHitJson(
            'page-1',
            'drive-1',
            'SitePage-Example-1.aspx',
            'https://home3925.sharepoint.com/sites/DemoSite/SitePages/SitePage-Example-1.aspx'
        );
        String fileHit = hitJson(
            'item-5',
            'drive-1',
            'testdocument.docx',
            'Jane Doe',
            '2026-06-01T10:00:00Z',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx'
        );
        Test.setMock(
            HttpCalloutMock.class,
            new GraphApiMock(
                responseWithHits(new List<String>{ sitePageHit, fileHit }),
                new Map<String, Map<String, String>>()
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            null,
            null,
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(
            1,
            results.size(),
            'Expected the .aspx Site Page hit to be filtered out'
        );
        System.assertEquals('testdocument.docx', results[0].title);
    }

    @IsTest
    static void searchDocumentsBuildsExpectedQueryTextAndOmitsFieldsFromTheSearchRequest() {
        GraphApiMock mock = new GraphApiMock(
            responseWithHits(new List<String>()),
            new Map<String, Map<String, String>>()
        );
        Test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        SharePointSearchController.searchDocuments(
            'test',
            'Contract',
            '12345',
            '00654321',
            'Jane Doe'
        );
        Test.stopTest();

        Map<String, Object> body = (Map<String, Object>) JSON.deserializeUntyped(
            mock.capturedSearchRequest.getBody()
        );
        List<Object> requests = (List<Object>) body.get('requests');
        Map<String, Object> searchRequest = (Map<String, Object>) requests[0];
        Map<String, Object> query = (Map<String, Object>) searchRequest.get(
            'query'
        );
        String queryText = (String) query.get('queryString');

        System.assertEquals('POST', mock.capturedSearchRequest.getMethod());
        System.assert(
            queryText.contains('filename:"test*"'),
            'Expected a prefix-wildcard filename clause so "test" matches "testdocument.docx"'
        );
        System.assert(
            queryText.contains('DocumentType:"Contract"'),
            'Expected an exact-match DocumentType clause'
        );
        System.assert(
            queryText.contains('ProjectID:"12345"'),
            'Expected an exact-match ProjectID clause'
        );
        System.assert(
            queryText.contains('OpportunityID:"00654321"'),
            'Expected an exact-match OpportunityID clause'
        );
        System.assertNotEquals(
            null,
            searchRequest.get('region'),
            'Graph requires a region on searches using application permissions'
        );
        System.assertEquals(
            null,
            searchRequest.get('fields'),
            'No "fields" array should be requested on the search call - it blanks driveItem search results'
        );
    }

    @IsTest
    static void searchDocumentsWithNoCriteriaReturnsAllRows() {
        String hit = hitJson(
            'item-3',
            'drive-1',
            'Any.docx',
            'Someone',
            '2026-01-01T00:00:00Z',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/Any.docx'
        );
        Test.setMock(
            HttpCalloutMock.class,
            new GraphApiMock(
                responseWithHits(new List<String>{ hit }),
                new Map<String, Map<String, String>>()
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            null,
            null,
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(
            1,
            results.size(),
            'A blank search should still return whatever Graph returns'
        );
    }

    @IsTest
    static void searchDocumentsWithZeroResultsReturnsEmptyList() {
        Test.setMock(
            HttpCalloutMock.class,
            new GraphApiMock(
                responseWithHits(new List<String>()),
                new Map<String, Map<String, String>>()
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'Nothing Matches',
            null,
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(
            0,
            results.size(),
            'Expected an empty list, not null, when there are no matches'
        );
    }

    @IsTest
    static void searchDocumentsThrowsAuraHandledExceptionOnCalloutError() {
        Test.setMock(
            HttpCalloutMock.class,
            new SearchErrorMock(401, '{"error":{"message":"boom"}}')
        );

        Test.startTest();
        try {
            SharePointSearchController.searchDocuments(
                'Doc',
                null,
                null,
                null,
                null
            );
            System.assert(
                false,
                'Expected an AuraHandledException to be thrown on a non-200 response'
            );
        } catch (AuraHandledException e) {
            System.assert(true);
        }
        Test.stopTest();
    }

    private class SearchErrorMock implements HttpCalloutMock {
        private final Integer statusCode;
        private final String body;

        SearchErrorMock(Integer statusCode, String body) {
            this.statusCode = statusCode;
            this.body = body;
        }

        public HttpResponse respond(HttpRequest req) {
            HttpResponse res = new HttpResponse();
            res.setStatusCode(statusCode);
            res.setStatus('Error');
            res.setBody(body);
            return res;
        }
    }

    @IsTest
    static void getDownloadUrlReturnsTheDirectDownloadLink() {
        Test.setMock(
            HttpCalloutMock.class,
            new DownloadUrlMock(
                200,
                '{"@microsoft.graph.downloadUrl":"https://download.sharepoint.com/direct/testdocument.docx"}'
            )
        );

        Test.startTest();
        String downloadUrl = SharePointSearchController.getDownloadUrl(
            'drive-1',
            'item-1'
        );
        Test.stopTest();

        System.assertEquals(
            'https://download.sharepoint.com/direct/testdocument.docx',
            downloadUrl
        );
    }

    @IsTest
    static void getDownloadUrlThrowsAuraHandledExceptionWhenIdsAreMissing() {
        Test.startTest();
        try {
            SharePointSearchController.getDownloadUrl(null, null);
            System.assert(
                false,
                'Expected an AuraHandledException when driveId/itemId are missing'
            );
        } catch (AuraHandledException e) {
            System.assert(true);
        }
        Test.stopTest();
    }

    @IsTest
    static void getDownloadUrlThrowsAuraHandledExceptionOnCalloutError() {
        Test.setMock(HttpCalloutMock.class, new DownloadUrlMock(404, '{}'));

        Test.startTest();
        try {
            SharePointSearchController.getDownloadUrl('drive-1', 'item-1');
            System.assert(
                false,
                'Expected an AuraHandledException on a non-200 response'
            );
        } catch (AuraHandledException e) {
            System.assert(true);
        }
        Test.stopTest();
    }
}
