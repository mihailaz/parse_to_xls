<?php

$urls = [
	'/catalog/umbro/',
	'/catalog/umbro/?PAGEN_1=2',
	'/catalog/umbro/?PAGEN_1=3',
	'/catalog/umbro/?PAGEN_1=4',
	'/catalog/umbro/?PAGEN_1=5',
	'/catalog/umbro/?PAGEN_1=6',
];

$pages = [];

$data = [];

if (!file_exists('./cache'))
	mkdir('./cache');

/**
 * @param string $url
 * @return string content
 * @throws Exception
 */
function getContent($url, $require = true)
{
	if (!$url){
		if ($require)
			throw new \Exception('Invalid url');
		else
			return null;
	}

	$url = 'http://www.proball.ru' . $url;
	$key = md5($url);
	$cache_file = "./cache/$key";

	if (file_exists($cache_file))
		return file_get_contents($cache_file);

	$content = file_get_contents($url);
	file_put_contents($cache_file, $content);

	return $content;
}

foreach ($urls as $url)
{
	$content = getContent($url);

	$domDoc  = new \DOMDocument('1.0', 'utf-8');
	$domDoc->strictErrorChecking = false;
	$success = @$domDoc->loadHTML($content);

	$errors = libxml_get_errors();
	if ($errors)
		throw new \Exception('Invalid html document');
	if (!$success)
		throw new \Exception('Error parsing document');

	$xpath = new \DOMXPath($domDoc);

	// xpath to item link
	$nodeList = $xpath->query("descendant-or-self::*[@id = 'catalog_panel']/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' items_list ')]/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' item ')]");

	/**
	 * @var \DOMElement $a
	 */
	foreach ($nodeList as $a)
	{
		$href = $a->getAttribute('href');

		if (!$href)
			continue;

		$pages[] = $href;
	}
}

foreach ($pages as $url)
{
	$content = getContent($url);

	$domDoc  = new \DOMDocument('1.0', 'utf-8');
	$domDoc->strictErrorChecking = false;
	$success = @$domDoc->loadHTML($content);

	$errors = libxml_get_errors();
	if ($errors)
		throw new \Exception('Invalid html document');
	if (!$success)
		throw new \Exception('Error parsing document');

	$xpath = new \DOMXPath($domDoc);
	$nodeTitle = $xpath->query("descendant-or-self::*[@id = 'WorkForm']/descendant::h1")->item(0);

	if (!$nodeTitle)
		throw new \Exception('title not found');

	$title = trim($nodeTitle->textContent);

	$nodeCode = $xpath->query("descendant-or-self::*[@id = 'article']")->item(0);

	if (!$nodeCode)
		throw new \Exception('code not found');

	$code = trim($nodeCode->nodeValue);

	$nodePrice = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' price_panel ')]/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' price ')]")->item(0);

	if (!$nodePrice)
		throw new \Exception('price not found');

	$price = trim($nodePrice->textContent);

	$price = explode(' ', $price);
	$currency = $price[1];
	$price = $price[0];

	if (!$price || !$currency)
		throw new \Exception('Invalid price');

	$nodeDesc = $xpath->query("descendant-or-self::*[@id = 'tabs-1']/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' info ')]")->item(0);

	if (!$nodeDesc)
		throw new \Exception('description not found');

	$desc = trim($nodeDesc->textContent);

	$nodesSpec = $xpath->query("descendant-or-self::*[@id = 'tabs-2']/descendant::table/descendant::tr/*");

	if (!$nodesSpec || !$nodesSpec->length)
		throw new \Exception('Spec not found');

	$spec = [];
	$last_key = null;

	/**
	 * @var \DOMElement $s
	 */
	foreach ($nodesSpec as $s)
	{
		$s = trim($s->textContent);

		if ($last_key)
		{
			$spec[$last_key] = $s;
			$last_key = null;
		}
		else
			$last_key = $s;
	}

	$nodeImg = $xpath->query("descendant-or-self::*[@id = 'WorkForm']/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' big_img ')]/descendant::img")->item(0);

	if (!$nodeImg)
		throw new \Exception('image not found');

	$img = trim($nodeImg->getAttribute('src'));

	$file_name = end(explode('/', $img));
	$img_src = getContent($img, false);

	if ($img_src)
		file_put_contents($file_name, $img_src);

	$data[] = [
		'title' => $title,
		'code' => $code,
		'price' => $price,
		'currency' => $currency,
		'description' => $desc,
		'spec' => $spec,
		'img' => $file_name,
	];
}

if (!$data)
	throw new \Exception('Empty data');

$table = '';

foreach ($data as $d)
{
	$row = '<Row>';

	foreach ($d as $k => $v)
	{
		if (is_array($v))
		{
			$a = [];
			foreach ($v as $ks => $vs)
			{
				$a[] = $ks . ': ' . $vs;
			}

			$v = implode("\r\n", $a);
		}

		$row .= <<<TABLE
<Cell><Data ss:Type="String">{$v}</Data></Cell>
TABLE;
	}

	$row .= '</Row>';
	$table .= $row;
}


$tmpl = <<<TMPL
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
	<Styles>
  		<Style ss:ID="bold">
			<Font ss:Bold="1"/>
		</Style>
 	</Styles>
	<Worksheet ss:Name="WorksheetName">
		<Table>
			$table
		</Table>
	</Worksheet>
</Workbook>
TMPL;

file_put_contents('export.xls', $tmpl);
