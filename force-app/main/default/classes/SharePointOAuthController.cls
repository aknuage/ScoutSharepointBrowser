public with sharing class SharePointOauthController {
    /** Called automatically by the VF page’s action attribute */
    public PageReference init() {
    String code = ApexPages.currentPage().getParameters().get('code');
    if (String.isNotBlank(code)) {
        handleTokenExchange(code);
    }
    // **do not** do any PageReference redirect here!
    return null;
    }

    /** Handles the code→token exchange with Microsoft */
    private void handleTokenExchange(String code) {
        // load your Azure creds & tenant from your custom metadata
        Sharepoint_Config__mdt cfg = Sharepoint_Config__mdt.getInstance('Scout_Site');
        if (cfg == null) {
            throw new AuraHandledException('SharePoint config not found.');
        }

        String clientId     = cfg.Client_ID__c;
        String clientSecret = cfg.Client_Secret__c;
        String tenantId     = cfg.Tenant_ID__c;
        String redirectUri  = URL.getOrgDomainUrl().toExternalForm() + '/apex/SharePointCallback';

        // build the token‐exchange request
        HttpRequest req = new HttpRequest();
        req.setEndpoint('https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token');
        req.setMethod('POST');
        req.setHeader('Content-Type', 'application/x-www-form-urlencoded');
        req.setBody(
            'grant_type=authorization_code'
            + '&code=' + EncodingUtil.urlEncode(code, 'UTF-8')
            + '&redirect_uri=' + EncodingUtil.urlEncode(redirectUri, 'UTF-8')
            + '&client_id=' + clientId
            + '&client_secret=' + EncodingUtil.urlEncode(clientSecret, 'UTF-8')
        );

        Http http = new Http();
        HttpResponse res = http.send(req);

        if (res.getStatusCode() != 200) {
            throw new AuraHandledException('Token exchange failed: ' + res.getBody());
        }

        // parse and upsert the tokens
        Map<String, Object> json = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
        String accessToken  = (String) json.get('access_token');
        String refreshToken = (String) json.get('refresh_token');
        Integer expiresIn   = (Integer) json.get('expires_in');

        upsertSharePointToken(accessToken, refreshToken, expiresIn);
    }

    /** Creates or updates the single SharePoint_Token__c record for the current user */
    private void upsertSharePointToken(String accessToken, String refreshToken, Integer expiresInSeconds) {
        Id thisUserId = UserInfo.getUserId();
        DateTime expiresAt = DateTime.now().addSeconds(expiresInSeconds);

        // Query into a List
        List<SharePoint_Token__c> listTokens = [
            SELECT Id
            FROM SharePoint_Token__c
            WHERE User_Id__c = :thisUserId
            LIMIT 1
        ];

        SharePoint_Token__c token;
        if (listTokens.isEmpty()) {
            // no existing record → create new
            token = new SharePoint_Token__c(
                User_Id__c = thisUserId
            );
        } else {
            // use the existing one
            token = listTokens[0];
        }

        token.Access_Token__c  = accessToken;
        token.Refresh_Token__c = refreshToken;
        token.Expires_At__c    = expiresAt;
        upsert token;
    }



    @AuraEnabled
    public static Boolean checkSharePointToken() {
        List<SharePoint_Token__c> tokenList = [
            SELECT Access_Token__c, Expires_At__c
            FROM SharePoint_Token__c
            WHERE User_Id__c = :UserInfo.getUserId()
        ];

        if (tokenList.isEmpty()) {
            throw new AuraHandledException('No token record found for this user.');
        }

        SharePoint_Token__c token = tokenList[0];
        return token.Access_Token__c != null && token.Expires_At__c > System.now();
    }

    /** 
     * Returns the current user’s SharePoint access token.
     * @return the Bearer token string
     */
    public static String getAccessTokenForCurrentUser() {
        List<SharePoint_Token__c> tokens = [
            SELECT Access_Token__c, Expires_At__c
            FROM SharePoint_Token__c
            WHERE User_Id__c = :UserInfo.getUserId()
            LIMIT 1
        ];
        if (tokens.isEmpty()) {
            throw new AuraHandledException('No token record found for this user.');
        }
        SharePoint_Token__c token = tokens[0];
        if (token.Expires_At__c == null || token.Expires_At__c < System.now()) {
            throw new AuraHandledException('Access token has expired.');
        }
        return token.Access_Token__c;
    }

}
