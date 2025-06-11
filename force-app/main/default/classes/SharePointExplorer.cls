public class SharePointExplorer {
    public static void explore(String driveId, String path) {
        String endpoint;

        if (String.isBlank(path)) {
            endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/root/children';
        } else {
            String encodedPath = EncodingUtil.urlEncode(path, 'UTF-8').replace('+', '%20');
            endpoint = 'callout:MicrosoftGraph/v1.0/drives/' + driveId + '/root:/' + encodedPath + ':/children';
        }

        System.debug('📂 Exploring: ' + (path == null ? '[root]' : path) + ' → ' + endpoint);

        HttpRequest req = new HttpRequest();
        req.setEndpoint(endpoint);
        req.setMethod('GET');
        req.setHeader('Accept', 'application/json');

        Http http = new Http();
        HttpResponse res = http.send(req);

        if (res.getStatusCode() == 200) {
            Map<String, Object> json = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
            List<Object> items = (List<Object>) json.get('value');

            for (Object o : items) {
                Map<String, Object> item = (Map<String, Object>) o;
                String name = (String) item.get('name');
                Boolean isFolder = item.containsKey('folder');

                String fullPath = String.isBlank(path) ? name : path + '/' + name;
                System.debug((isFolder ? '📁' : '📄') + ' ' + fullPath);

                // Recursive call for folders
                if (isFolder) {
                    explore(driveId, fullPath);
                }
            }
        } else {
            System.debug('❌ Error: ' + res.getStatusCode() + ' → ' + res.getBody());
        }
    }
}