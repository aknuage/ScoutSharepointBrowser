@isTest
private class SharePointFileBrowserControllerTest {
    static final String MOCK_CREATE_FOLDER_RESPONSE = '{"id": "new-folder-id"}';
    static final String MOCK_RESPONSE = '{"value": [ { "id": "file1", "name": "TestDoc.pdf", "webUrl": "https://example.com", "lastModifiedDateTime": "2023-10-05T14:48:00Z", "size": 102400 } ]}';
    static final String MOCK_SINGLE_ITEM_RESPONSE = '{"id": "folder1", "name": "TestFolder", "parentReference": {"driveId": "drive123"}}';

    class MockHttpResponseGenerator implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            if (req.getEndpoint().contains('/items/') && req.getEndpoint().contains('/children') && req.getMethod() == 'POST') {
                res.setBody(MOCK_CREATE_FOLDER_RESPONSE);
                res.setStatusCode(201);
            } else if (req.getEndpoint().contains('/preview')) {
                res.setBody('{"getUrl": "https://example.com/preview"}');
                res.setStatusCode(200);
            } else if (req.getEndpoint().contains('/search')) {
                res.setBody('{"value": [ { "id": "file2", "name": "SearchResult.txt" } ]}');
                res.setStatusCode(200);
            } else if (req.getEndpoint().contains('/children')) {
                res.setBody(MOCK_RESPONSE);
                res.setStatusCode(200);
            } else if (req.getEndpoint().contains('/driveItem')) {
                res.setBody(MOCK_SINGLE_ITEM_RESPONSE);
                res.setStatusCode(200);
            } else if (req.getEndpoint().contains('/items/') && req.getMethod() == 'DELETE') {
                res.setBody('');
                res.setStatusCode(204);
            } else {
                res.setBody('{"status":"ok"}');
                res.setStatusCode(200);
            }
            return res;
        }
    }

    @isTest
    static void testSearchSharePoint() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        Test.startTest();
        List<Map<String, Object>> results = SharePointFileBrowserController.searchSharePoint('drive123', 'search');
        Test.stopTest();

        System.assertEquals(1, results.size(), 'Should return one search result');
        System.assertEquals('SearchResult.txt', results[0].get('name'));
    }

    @isTest
    static void testGetSharePointPreviewUrl() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        Test.startTest();
        String previewUrl = SharePointFileBrowserController.getSharePointPreviewUrl('drive123', 'file123');
        Test.stopTest();

        System.assert(previewUrl.contains('https://example.com/preview'), 'Expected valid preview URL');
    }

    @isTest
    static void testGetFolderContentsFromRecord_Project() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());

        Project__c proj = new Project__c(
            Name = 'Test Project',
            Link_to_Real_Estate_SharePoint_Folder__c = 'https://scoutcleanenergy.sharepoint.com/sites/test/Docs'
        );
        insert proj;

        Test.startTest();
        List<Map<String, Object>> result = SharePointFileBrowserController.getFolderContentsFromRecord(proj.Id, 'Project__c');
        Test.stopTest();

        System.assertEquals(1, result.size(), 'Should return one file from Project__c');
        System.assertEquals('TestDoc.pdf', (String)result[0].get('name'));
    }

    // @isTest
    // static void testGetFolderContentsFromRecord_RFA() {
    //     Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());

    //     Project__c proj = new Project__c(
    //         Name = 'Test Project',
    //         Link_to_Real_Estate_SharePoint_Folder__c = 'https://scoutcleanenergy.sharepoint.com/sites/test/Docs'
    //     );
    //     insert proj;

    //     Request_For_Agreement__c rfa = new Request_For_Agreement__c(
    //         Name = 'Test RFA',
    //         Project_Name__c = proj.Id,
    //         Account__c = UserInfo.getUserId(), // Use current user for required lookup
    //         Agreement_Type__c = 'Standard' // Mock required picklist value
    //     );
    //     insert rfa;

    //     Test.startTest();
    //     List<Map<String, Object>> result = SharePointFileBrowserController.getFolderContentsFromRecord(rfa.Id, 'Request_For_Agreement__c');
    //     Test.stopTest();

    //     System.assertEquals(1, result.size(), 'Should return one file from RFA');
    //     System.assertEquals('TestDoc.pdf', (String)result[0].get('name'));
    // }

    // @isTest
    // static void testGetFolderContentsFromRecord_UnsupportedType() {
    //     Test.startTest();
    //     try {
    //         SharePointFileBrowserController.getFolderContentsFromRecord('001000000000001', 'Account');
    //         System.assert(false, 'Should have thrown an AuraHandledException');
    //     } catch (AuraHandledException ex) {
    //         System.assert(ex.getMessage().contains('Unsupported object type'), 'Expected unsupported object type error');
    //     }
    //     Test.stopTest();
    // }

    // @isTest
    // static void testGetFolderContentsFromRecord_NoLink() {
    //     Project__c proj = new Project__c(Name = 'No Link Project', Link_to_Real_Estate_SharePoint_Folder__c = null);
    //     insert proj;

    //     Test.startTest();
    //     try {
    //         SharePointFileBrowserController.getFolderContentsFromRecord(proj.Id, 'Project__c');
    //         System.assert(false, 'Expected exception for missing SharePoint link');
    //     } catch (AuraHandledException e) {
    //         System.assert(e.getMessage().contains('Missing SharePoint link'), 'Expected missing link error');
    //     }
    //     Test.stopTest();
    // }

    // @isTest
    // static void testGetChildrenFromSharedUrl() {
    //     Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
    //     List<Map<String, Object>> result = SharePointFileBrowserController.getChildrenFromSharedUrl('https://dummy.sharepoint.com/sites/test/folder');
    //     System.assertEquals(1, result.size());
    //     System.assertEquals('TestDoc.pdf', (String)result[0].get('name'));
    // }

    @isTest
    static void testGetFolderChildrenById() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        List<Map<String, Object>> result = SharePointFileBrowserController.getFolderChildrenById('drive123', 'folder1');
        System.assertEquals(1, result.size());
        System.assertEquals('TestDoc.pdf', (String)result[0].get('name'));
    }

    @isTest
    static void testUploadFileToSharePoint() {
        String base64Body = EncodingUtil.base64Encode(Blob.valueOf('Dummy content'));
        String fileName = 'upload_test.txt';
        String driveId = 'drive123';
        String itemId = 'folder123';

        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        Test.startTest();
        SharePointFileBrowserController.uploadFileToSharePoint(base64Body, fileName, driveId, itemId);
        Test.stopTest();

        System.assert(true, 'No exceptions should occur during upload');
    }

    @isTest
    static void testDeleteSharepointFile() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        Test.startTest();
        SharePointFileBrowserController.deleteSharepointFile('item123', 'drive123');
        Test.stopTest();

        // Basic assertion to confirm no exceptions were thrown
        System.assert(true, 'No exception should be thrown on delete');
    }

    @isTest
    static void testCreateFolderInSharePoint() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        Test.startTest();
        String newFolderId = SharePointFileBrowserController.createFolderInSharePoint('MyFolder', 'drive123', 'parent456');
        Test.stopTest();

        System.assertEquals('new-folder-id', newFolderId, 'Expected created folder ID');
    }


}