/**
 * Exposes a method to browse the contents of a SharePoint document library via Microsoft Graph API.
 * Uses a Named Credential and custom metadata to access a specific site and drive root.
 * @author AKelly @ NuAge
 * @date March 2025
 */
public with sharing class SharePointFileBrowserController {

    // SharePoint Site ID loaded from Custom Metadata
    private static final String SITE_ID = SharePoint_Config__mdt.getInstance('Scout_Site').Site_ID__c;
    private static String driveId;

    /**
     * Fetches a preview URL for a SharePoint file.
     * @param driveId - The drive ID where the file resides.
     * @param itemId - The item ID of the file.
     * @return String - A preview URL for embedding.
     * @author AKelly @ NuAge
     * @date April 2025
     */
    @AuraEnabled
    public static String getSharePointPreviewUrl(String driveId, String itemId) {
        if (String.isBlank(driveId) || String.isBlank(itemId)) {
            throw new AuraHandledException('Missing driveId or itemId.');
        }

        String endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/items/' + itemId + '/preview';

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('POST');
        req.setHeader('Accept', 'application/json');
        req.setHeader('Content-Type', 'application/json');
        req.setBody('{}'); // 🟢 Required – JSON body for POST

        Http http = new Http();
        HttpResponse res = http.send(req);

        if (res.getStatusCode() != 200) {
            System.debug('Error fetching preview URL: ' + res.getBody());
            throw new AuraHandledException('Error fetching preview URL: ' + res.getBody());
        }

        Map<String, Object> json = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
        String previewUrl = (String) json.get('getUrl');
        System.debug('previewUrl: ' + previewUrl);
        if (String.isBlank(previewUrl)) {
            throw new AuraHandledException('No preview URL returned.');
        }

        return previewUrl + '&nb=true';
    }

    /**
     * Determines which SharePoint link to use based on the object type.
     * @param recordId - The Id of the source record (Project or RFA)
     * @param objectApiName - The API name of the object (e.g. 'Project__c' or 'Request_For_Agreement__c')
     * @return List<Map<String, Object>> - Folder contents.
     * @author AKelly @ NuAge
     * @date April 2025
     */
    @AuraEnabled
    public static List<Map<String, Object>> getFolderContentsFromRecord(Id recordId, String objectApiName) {
        String sharePointUrl;

        if (objectApiName == 'Project__c') {
            sharePointUrl = [
                SELECT Link_to_Real_Estate_SharePoint_Folder__c 
                FROM Project__c 
                WHERE Id = :recordId
            ].Link_to_Real_Estate_SharePoint_Folder__c;
        } 
        else if (objectApiName == 'Request_For_Agreement__c') {
            sharePointUrl = [
                SELECT Project_Name__r.Link_to_Real_Estate_SharePoint_Folder__c 
                FROM Request_For_Agreement__c 
                WHERE Id = :recordId
            ].Project_Name__r.Link_to_Real_Estate_SharePoint_Folder__c;
        } 
        else {
            throw new AuraHandledException('Unsupported object type: ' + objectApiName);
        }

        if (String.isBlank(sharePointUrl)) {
            throw new AuraHandledException('Missing SharePoint link.');
        }

        return getChildrenFromSharedUrl(sharePointUrl);
    }


    /**
     * Returns child items of a folder specified by its shared URL.
     * @param sharedUrl - SharePoint shared folder URL.
     * @return List<Map<String, Object>> - Child items.
     * @author AKelly @ NuAge
     * @date April 2025
     */
    public static List<Map<String, Object>> getChildrenFromSharedUrl(String sharedUrl) {
        String encodedUrl = EncodingUtil.base64Encode(Blob.valueOf(sharedUrl))
            .replace('+', '-')
            .replace('/', '_')
            .replace('=', '');
        String endpoint = 'callout:MicrosoftGraph/v1.0/shares/u!' + encodedUrl + '/driveItem';

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('GET');
        req.setHeader('Accept', 'application/json');

        Http http = new Http();
        HttpResponse res = http.send(req);
        if (res.getStatusCode() != 200) {
            throw new AuraHandledException('Failed to resolve folder from SharePoint URL: ' + res.getBody());
        }

        Map<String, Object> data = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
        Map<String, Object> parentRef = (Map<String, Object>) data.get('parentReference');
        String driveId = (String) parentRef.get('driveId');
        String itemId = (String) data.get('id');

        String childrenEndpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/items/' + itemId + '/children';
        HttpRequest childReq = new HttpRequest();
        childReq.setEndpoint(childrenEndpoint);
        childReq.setMethod('GET');
        childReq.setHeader('Accept', 'application/json');

        HttpResponse childRes = http.send(childReq);
        if (childRes.getStatusCode() != 200) {
            throw new AuraHandledException('Failed to fetch folder children: ' + childRes.getBody());
        }

        Map<String, Object> childrenData = (Map<String, Object>) JSON.deserializeUntyped(childRes.getBody());
        List<Object> rawList = (List<Object>) childrenData.get('value');

        List<Map<String, Object>> result = new List<Map<String, Object>>();
        for (Object o : rawList) {
            result.add((Map<String, Object>) o);
        }
        return result;
    }

    /**
     * Fetches children of a folder by drive and item Ids.
     * @param driveId - SharePoint drive ID.
     * @param itemId - SharePoint folder item ID.
     * @return List<Map<String, Object>> - Children.
     * @author AKelly @ NuAge
     * @date April 2025
    */
    @AuraEnabled
    public static List<Map<String, Object>> getFolderChildrenById(String driveId, String itemId) {
        String endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/items/' + itemId + '/children';

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('GET');
        req.setHeader('Accept', 'application/json');

        Http http = new Http();
        HttpResponse res = http.send(req);

        if (res.getStatusCode() != 200) {
            throw new AuraHandledException('Failed to fetch folder children by ID: ' + res.getBody());
        }

        Map<String, Object> childrenData = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
        List<Object> rawList = (List<Object>) childrenData.get('value');

        List<Map<String, Object>> result = new List<Map<String, Object>>();
        for (Object o : rawList) {
            result.add((Map<String, Object>) o);
        }
        return result;
    }

    /**
     * Uploads a file directly to a SharePoint folder.
     * @param base64Body - The file body encoded in base64.
     * @param fileName - Name of the file.
     * @param driveId - SharePoint drive ID.
     * @param itemId - Target folder ID in SharePoint.
     * @author AKelly @ NuAge
     * @date April 2025
    */
    @AuraEnabled
    public static void uploadFileToSharePoint(String base64Body, String fileName, String driveId, String itemId) {
        if (String.isBlank(base64Body) || String.isBlank(fileName)) {
            throw new AuraHandledException('Missing file data or filename.');
        }
        // Decode
        Blob fileData = EncodingUtil.base64Decode(base64Body);

        // Fetch the user token
        String userToken = SharePointOauthController.getAccessTokenForCurrentUser();
        String bearer = 'Bearer ' + userToken;

        // Direct Graph call
        String endpoint = 'https://graph.microsoft.com/v1.0/drives/'
            + driveId + '/items/' + itemId + ':/' + EncodingUtil.urlEncode(fileName,'UTF-8') + ':/content';

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('PUT');
        req.setHeader('Authorization', bearer);
        req.setHeader('Content-Type', 'application/octet-stream');
        req.setBodyAsBlob(fileData);

        Http http = new Http();
        HttpResponse res = http.send(req);
        if (res.getStatusCode() != 201 && res.getStatusCode() != 200) {
            throw new AuraHandledException('Error uploading file: ' + res.getBody());
        }
    }

    /**
     * Searches a SharePoint document library using Microsoft Graph API.
     * Filters results to only include items whose names contain the search term (case-insensitive).
     * Requires a valid drive ID and a non-blank search string.
     * @param driveId - SharePoint drive ID to search within.
     * @param searchText - The text to search for in file or folder names.
     * @return List<Map<String, Object>> - Matching SharePoint items.
     * @author AKelly @ NuAge
     * @date April 2025
     */
    @AuraEnabled
    public static List<Map<String, Object>> searchSharePoint(String driveId, String searchText) {
        if (String.isBlank(driveId) || String.isBlank(searchText)) {
            throw new AuraHandledException('Missing driveId or searchText.');
        }

        String endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/root/search(q=\'' + EncodingUtil.urlEncode(searchText, 'UTF-8') + '\')';

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('GET');
        req.setHeader('Accept', 'application/json');

        Http http = new Http();
        HttpResponse res = http.send(req);

        if (res.getStatusCode() != 200) {
            throw new AuraHandledException('Search failed: ' + res.getBody());
        }

        Map<String, Object> json = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
        List<Object> items = (List<Object>) json.get('value');

        List<Map<String, Object>> results = new List<Map<String, Object>>();
        for (Object obj : items) {
            Map<String, Object> item = (Map<String, Object>) obj;
            String name = (String) item.get('name');
            if (name != null && name.toLowerCase().contains(searchText.toLowerCase())) {
                results.add(item);
            }
        }

        return results;
    }

    @AuraEnabled 
    public static void deleteSharepointFile(String itemId, String driveId)
    {
        if (String.isBlank(itemId) || String.isBlank(driveId)) {
            throw new AuraHandledException('Missing driveId or itemId.');
        }
    
        String endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/items/' + itemId;
    
        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('DELETE');
        req.setHeader('Accept', 'application/json');
    
        Http http = new Http();
        HttpResponse res = http.send(req);
    
        if (res.getStatusCode() != 204 && res.getStatusCode() != 200) {
            System.debug('Error deleting item: ' + res.getBody());
            throw new AuraHandledException('Error deleting item: ' + res.getBody());
        }
    }
    

    /**
     * Creates a new folder in a specified SharePoint drive and parent folder.
     * @param folderName - Name of the folder to create.
     * @param driveId - SharePoint Drive ID.
     * @param parentItemId - ID of the parent folder where new folder will be created.
     * @return String - ID of the newly created folder.
     * @author AKelly @ NuAge
     * @date April 2025
     */
    @AuraEnabled
    public static String createFolderInSharePoint(String folderName, String driveId, String parentItemId) {
        if (String.isBlank(folderName) || String.isBlank(driveId) || String.isBlank(parentItemId)) {
            throw new AuraHandledException('Missing required information to create folder.');
        }

        String endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/items/' + parentItemId + '/children';

        Map<String, Object> body = new Map<String, Object>{
            'name' => folderName,
            'folder' => new Map<String, Object>{},
            '@microsoft.graph.conflictBehavior' => 'rename'
        };

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('POST');
        req.setHeader('Content-Type', 'application/json');
        req.setBody(JSON.serialize(body));

        Http http = new Http();
        HttpResponse res = http.send(req);

        if (res.getStatusCode() != 201 && res.getStatusCode() != 200) {
            throw new AuraHandledException('Error creating folder: ' + res.getBody());
        }

        Map<String, Object> result = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
        return (String) result.get('id'); // ID of the newly created folder
    }


    @AuraEnabled
    public static String initiateAuthFlow() {
    Sharepoint_Config__mdt cfg = Sharepoint_Config__mdt.getInstance('Scout_Site');
    String clientId     = cfg.Client_ID__c;
    String tenantId     = cfg.Tenant_ID__c;
    String redirectUri  = URL.getOrgDomainUrl().toExternalForm() + '/apex/SharePointCallback';
    return 'https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/authorize'
        + '?client_id='     + clientId
        + '&response_type=code'
        + '&redirect_uri='   + EncodingUtil.urlEncode(redirectUri,'UTF-8')
        + '&response_mode=query'
        + '&scope='          + EncodingUtil.urlEncode('offline_access Files.ReadWrite.All','UTF-8')
        + '&state='          + UserInfo.getUserId();
    }

 
}