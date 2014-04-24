<?php

/**
 * Helper for accessing google drive spreadsheet data. (Because the `google/apiclient` is missing a Google_Service_Spreadsheet component.
 *
 * @link https://developers.google.com/google-apps/spreadsheets/
 */
class SpreadsheetService extends Sledgehammer\Object {

    /**
     * @var Google_Client
     */
    private $client;

    /**
     * @param Google_Client $client
     */
    public function __construct($client) {
        $this->client = $client;
    }

    /**
     * List available spreadsheets.
     *
     * @return array
     */
    public function getFiles() {
        $response = $this->call('https://spreadsheets.google.com/feeds/spreadsheets/private/full'); //?v=3.0
        $files = array();

        $xml = new SimpleXMLElement($response->getResponseBody());
        foreach ($xml->entry as $entry) {
            $files[] = array(
                'id' => (string) preg_replace('/^.+full\//', '', $entry->id),
                'title' => (string) $entry->title,
                'updated' => (string) $entry->updated,
                'author' => array(
                    'name' => (string) $entry->author->name,
                    'email' => (string) $entry->author->email,
                ),
                'self' => (string) $xml->id,
            );
        }
        return $files;
    }

    /**
     * Get file info and ids for the pages
     *
     * @param string $id  The spreadsheet ID
     * @return array
     */
    public function getFile($id) {
        $response = $this->call('https://spreadsheets.google.com/feeds/worksheets/' . $id . '/private/values');
        $xml = new SimpleXMLElement($response->getResponseBody());
        $file = array(
            'id' => (string) preg_replace('/^.+full\//', '', $xml->id),
            'title' => (string) $xml->title,
            'updated' => (string) $xml->updated,
            'author' => array(
                'name' => (string) $xml->author->name,
                'email' => (string) $xml->author->email,
            ),
            'pages' => array(),
            'self' => (string) $xml->id,
        );
        foreach ($xml->entry as $entry) {
            $pageId = preg_replace('/^.+values\//', '', $entry->id);
            $file['pages'][] = $pageId;
        }
        return $file;
    }

    /**
     * Retrieve page info including all rows and columns.
     *
     * @param string $id Spreadsheet ID
     * @param string $pageId Page/Tab ID
     * @return array
     */
    public function getPage($id, $pageId = 'od6', $options = array()) {
        // Merge default options
        $default = array(
            'parseHeaders' => false,
            'defaultValue' => ''
        );
        $options = array_merge($default, $options);
        if (count($options) !== count($default)) {
            trigger_error('Option: "' . implode('" and "', array_keys(array_diff_key($options, $default))) . '" is invalid', E_USER_WARNING);
        }

        $response = $this->call('https://spreadsheets.google.com/feeds/cells/' . $id . '/' . $pageId . '/private/values');
        $xml = new SimpleXMLElement($response->getResponseBody());
        $page = array(
            'id' => $pageId,
            'title' => (string) $xml->title,
            'self' => (string) $xml->id,
            'rows' => array()
        );
        $cells = array();

        foreach ($xml->entry as $entry) {
            $x = $this->cellX($entry->title);
            $y = $this->cellY($entry->title);
            $cells[$y][$x] = (string) $entry->content;
        }
        $page['rows'] = $cells;
        if ($options['parseHeaders']) {
            $rows = array();
            $headers = array_shift($cells);

            foreach ($cells as $y => $values) {
                $row = array();
                foreach ($headers as $x => $header) {
                    if (array_key_exists($x, $values)) {
                        $row[$header] = $values[$x];
                    } else {
                        $row[$header] = $options['defaultValue'];
                    }
                }
                $rows[] = $row;
            }
            $page['rows'] = $rows;
        }
        return $page;
    }

    /**
     * Convert cellname to X coordinate.
     *
     * @param string $title A1, C4 etc
     * @return int A1 = 0, B4 = 3
     */
    private function cellX($title) {
        $letters = preg_replace('/[0-9]*$/', '', $title);
        $count = strlen($letters);
        $n = 0;
        for ($i = 0; $i < $count; $i++) {
            $n = $n * 26 + ord($letters[$i]) - 0x40;
        }
        return $n - 1;
    }

    /**
     * Convert cellname to Y coordinate.
     *
     * @param string $title A1, C4 etc
     * @return int A1 = 0, B4 = 1
     */
    private function cellY($title) {
        return preg_replace('/^[A-Z]*/i', '', $title) - 1;
    }

    /**
     * Perform an authenticated api call
     *
     * @param string $url
     * @return Google_Http_Request
     */
    public function call($url) {
        $request = new Google_Http_Request($url);
        $this->client->getAuth()->sign($request);
        return $this->client->getIo()->makeRequest($request); /* @var $response Google_Http_Request */
    }

}
