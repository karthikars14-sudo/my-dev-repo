@IsTest
private class SharePointSearchControllerTest {
    private class SearchMock implements HttpCalloutMock {
        private final Integer statusCode;
        private final String body;
        private HttpRequest capturedRequest;

        SearchMock(Integer statusCode, String body) {
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

    private static String responseWithHits(List<String> hits) {
        return '{"value":[{"searchTerms":[],"hitsContainers":[{"total":' +
            hits.size() +
            ',"moreResultsAvailable":false,"hits":[' +
            String.join(hits, ',') +
            ']}]}]}';
    }

    @IsTest
    static void searchDocumentsReturnsMappedResultsIncludingDriveAndItemIds() {
        String hit = hitJson(
            'item-1',
            'drive-1',
            'testdocument.docx',
            'Jane Doe',
            '2026-06-01T10:00:00Z',
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx'
        );
        Test.setMock(
            HttpCalloutMock.class,
            new SearchMock(200, responseWithHits(new List<String>{ hit }))
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'test',
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
        System.assertEquals(
            'https://home3925.sharepoint.com/sites/DemoSite/Documents/testdocument.docx',
            results[0].fileUrl,
            'Search results should carry the webUrl - the direct download URL is fetched on demand, not eagerly'
        );
        System.assertEquals('drive-1', results[0].driveId);
        System.assertEquals('item-1', results[0].itemId);
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
            new SearchMock(
                200,
                responseWithHits(new List<String>{ folderHit, fileHit })
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'test',
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
        String sitePageHit = hitJson(
            'page-1',
            'drive-1',
            'SitePage-Example-1.aspx',
            'Someone',
            '2026-04-28T20:09:44Z',
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
            new SearchMock(
                200,
                responseWithHits(new List<String>{ sitePageHit, fileHit })
            )
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
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
    static void searchDocumentsUsesAPrefixWildcardOnTheFilenameProperty() {
        SearchMock mock = new SearchMock(
            200,
            responseWithHits(new List<String>())
        );
        Test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        SharePointSearchController.searchDocuments('test', null);
        Test.stopTest();

        Map<String, Object> body = (Map<String, Object>) JSON.deserializeUntyped(
            mock.capturedRequest.getBody()
        );
        List<Object> requests = (List<Object>) body.get('requests');
        Map<String, Object> searchRequest = (Map<String, Object>) requests[0];
        Map<String, Object> query = (Map<String, Object>) searchRequest.get(
            'query'
        );

        System.assertEquals('POST', mock.capturedRequest.getMethod());
        System.assert(
            ((String) query.get('queryString')).contains('filename:"test*"'),
            'Expected a prefix-wildcard filename clause so "test" matches "testdocument.docx"'
        );
        System.assertNotEquals(
            null,
            searchRequest.get('region'),
            'Graph requires a region on searches using application permissions'
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
            new SearchMock(200, responseWithHits(new List<String>{ hit }))
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
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
            new SearchMock(200, responseWithHits(new List<String>()))
        );

        Test.startTest();
        List<SharePointSearchController.SharePointDocumentResult> results = SharePointSearchController.searchDocuments(
            'Nothing Matches',
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
            new SearchMock(401, '{"error":{"message":"boom"}}')
        );

        Test.startTest();
        try {
            SharePointSearchController.searchDocuments('Doc', null);
            System.assert(
                false,
                'Expected an AuraHandledException to be thrown on a non-200 response'
            );
        } catch (AuraHandledException e) {
            System.assert(true);
        }
        Test.stopTest();
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
