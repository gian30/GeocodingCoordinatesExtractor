//
getApikey(){
      Properties prop = new Properties();
        InputStream input = null;
        input = new FileInputStream("config.properties");
        prop.load(input);
        String key;
        switch (currentApi) {
            case 1:
                key = prop.getProperty("APIKeyGoogle");
                break;
            case 2:
                key = prop.getProperty("APIKeyBing");
                break;
            case 3:
                key = prop.getProperty("APIKeyYandex");
                break;
            default:
                key = "none";
                break;
        }
        return key;
}