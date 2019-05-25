# Geocoding coordinates extractor

An application that extracts coordinates from the directions by reading Excel files.

## Download app
https://www.gianlucalv.com/projects/uploads/geocoding_extractor/geocoding-extractor.zip

## Getting Started

This app is used for coordinates extraction from Excel documents containing the directions of locations. It uses Google Geocoding API for getting information from. It processes the information containing coordinates and saves it in Excel format.

### Prerequisites

```
API Key Google
API Key Bing
NetBeans
Java
```

### Build project

1. Open NetBeans
2. Choose File > Open Project and select the project folder
3. Build the application: Choose Run > Clean and Build Project.

### Running Application

1. Go to the project folder
2. Go to dist/ and run Geocoding.jar

### Generating API Keys

Google Geocoding API:

1. Go to https://console.cloud.google.com/google/maps-apis/overview
2. Create or select a project.
3. Click Continue to enable the API.
4. Go to the Credentials page, to get your API key. 

- Restrictions:
    - 25,000 Requests per day.
- You can enable billing on https://developers.google.com/maps/documentation/geocoding/usage-limits 
to get more requests.


Bing Maps API:

1. Go to https://www.bingmapsportal.com/
2. Sign in with the Microsoft account or create a new one.
3. Select My keys under My Account.
4. Choose the option to create a new key.
5. Provide the required information and click the Create button.

- Restrictions:
    - 2 jobs in process at the same time.
    - 50 jobs in a 24 hour period.
- For upgrading to enterprise version visit https://www.microsoft.com/maps/contact.aspx


Yandex Maps Geocoder API:

- You don't need to use the API Key for free use of Yandex Geocoder API.
- Restrictions:
    - 25,000 Requests per day.
- You only need to use API Key in the commercial version, for more information visit: https://tech.yandex.com/maps/commercial/

### Using Application

First of all, you need to Add an API Key for the API you want to use.
Then choose an Excel document with the directions which you want to get the coordinates from.
After getting the coordinates you will be able to save them on your computer as an Excel document.
