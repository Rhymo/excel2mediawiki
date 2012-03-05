<?php

/***
Copyright (c) 2010 Shawn M. Douglas (shawndouglas.com)
Copyright (c) 2012 Thomas Gries (changes, additions)

Permission is hereby granted, free of charge, to any person
obtaining a copy of this software and associated documentation
files (the "Software"), to deal in the Software without
restriction, including without limitation the rights to use,
copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
***/

echo "<html>
<head>
<meta charset='UTF-8'>
<title>excel2wiki | Excel xls to MediaWiki copy and paste converter</title>
</head>
<body>
<h1>Copy & Paste Excel-to-Mediawiki Converter</h1>
<form action='' method='post'>
	<textarea name='data' rows='10' cols='50'></textarea>
	<br>
	<input type='submit' />
	<input type='checkbox' id='format-header' name='format-header' checked='checked'><label for='format-header'><small>format header</small></label>
</form>";

function nbspTrim( $str ) {
	$s = trim( $str );
	return ( $s === "" ) ? "&amp;nbsp;" : $s ;
}

if ( $_SERVER['REQUEST_METHOD'] != 'POST' ) {

	echo "<small><b>Instructions:</b><br>
<ul>
<li>Copy & paste cells from Excel and click submit. Paste results into your MediaWiki page.</li>
</ul>
<hr>
Last updated 2012-03-05. <a style='text-decoration:none; color:blue;' href=\"https://github.com/Wikinaut/excel2mediawiki\">Source code</a> available. This script is based on <a href=\"https://github.com/sdouglas/excel2wiki\">Excel2Wiki</a> by Shawn M. Douglas, 2010.</br>
";
} else {
	$outbuf = "<h2>Result</h2>\n<pre>\n{| class=\"wikitable sortable\"\n";

	$formatHeader = isset( $_POST['format-header'] );
	$lines = preg_split( "/\n/", $_POST['data'] );
	$n = sizeof( $lines );

	foreach ( $lines as $index => $line ) {

		$columns = preg_split( "/\t/", $line );

		if ( $index == 0 ) {
			foreach ( $columns as $column ) {
				$outbuf .= ( $formatHeader ? "! " : "| " ) . nbspTrim( $column ) . "\n";
			}
			$outbuf .= "|--\n";
		} else {
			$data = implode( " ||&amp;nbsp; ", $columns );
			if ( $index < ( $n - 1 ) && ( trim( $data ) != "" ) ) {
				$outbuf .= '|' . $data;
				$outbuf .= "|--\n";
			}
		}
	}

	$outbuf .= "|}</pre>";

	// cleaning up unnecessary stuff and beautifying the wiki-table elements
	$outbuf = preg_replace( "#\|&amp;nbsp; ([^|][^|])#", "| $1", $outbuf );
	$outbuf = preg_replace( "#\|&amp;nbsp;\s+\|#", "| &amp;nbsp; |", $outbuf );
	$outbuf = str_replace ( "| || &amp;nbsp; || &amp;nbsp; || &amp;nbsp; |--\n", "", $outbuf );
	echo $outbuf;

}

echo "</body></html>";
