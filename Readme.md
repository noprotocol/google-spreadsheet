#  Google Spreadsheet

PHP client for accessing the [google drive spreadsheet api](https://developers.google.com/google-apps/spreadsheets/).
(Because the [google/apiclient](https://developers.google.com/api-client-library/php/) is missing a Google_Service_Spreadsheet component.)


## Installation

```
composer require noprotocol/google-spreadsheet
```

## Usage example

```
$client = new Google_Client();
$client->setScopes('https://spreadsheets.google.com/feeds'); // when using ouath, ask for permissions
// other client setup code ...

$service = new SpreadsheetService($client);
$file = $service->getFile($id);
foreach ($file['pages'] as $pageId) {
	$page = $service->getPage($id, $pageId, array('parseHeaders' => true));
	foreach ($page['rows'] as $row) {
		// do stuff...
	}
};
```