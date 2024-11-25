# Load the HtmlAgilityPack assembly
Add-Type -Path "C:\Users\claud\AppData\Local\PackageManagement\NuGet\Packages\HtmlAgilityPack.1.11.71\lib\netstandard2.0\HtmlAgilityPack.dll"
$path = "C:\Users\claud\Documents\Consulting\hacking\"


# Open a CSV file to write to
$csvFilePath = $path + "patriot.csv"
$csvWriter = [System.IO.StreamWriter]::new($csvFilePath)

for ($i = 1; $i -lt 223; $i++) {

if (Test-Path -Path ($path + "patriot$i.html")) {
    # Load the HTML content from a file
    $htmlContent = Get-Content -Path ($path + "patriot$i.html") -Raw
} else {
    Write-Host "File patriot$i.html does not exist at $path."
    Exit
}

# Parse the HTML content using HtmlAgilityPack
$htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
$htmlDoc.LoadHtml($htmlContent)

# Find the table (assuming you know some identifier, like an id or class)
$table = $htmlDoc.DocumentNode.SelectSingleNode("//table[@id='T1']")


if($i -eq 1)
{
# Write the header row if available
$headerRow = $table.SelectSingleNode("//thead/tr")
$headers = $headerRow.SelectNodes("th")
if ($headers) {
    $headerValues = $headers | ForEach-Object { 
        $cell = $_.InnerText.Replace('"', '""').Replace("`r`n", " ").Replace("&nbsp;", " ").Trim() -replace "\s+" , " ";
        if($headers.IndexOf($_) -eq 3 -or $headers.IndexOf($_) -eq 5 -or $headers.IndexOf($_) -eq 6)
        {
        $cell = $cell -replace " ", '","'
        $cell
        }
        else {
            $cell
        }
    }
    $csvWriter.WriteLine('"Parcel ID","Street Number","Street","Owner","Built","Type","Total Value","Beds","Baths","Lot size","Fin area","LUC Description","NHood","Sale date","Sale price","Book Page"')
}
}
# Write the data rows
$dataRows = $table.SelectNodes("//tbody/tr[position()>1]")
foreach ($row in $dataRows) {
    $cells = $row.SelectNodes("td")
    
    $cellValues = $cells | ForEach-Object { 
        $cell = $_.InnerText.Replace('"', '""').Replace("`r`n", " ").Replace("&nbsp;", " ").Trim() -replace "\s+" , " ";
        
        if($cells.IndexOf($_) -eq 1 )
        {
            if($cell -match "^(\d+) ")
            {
                $cell = $cell -replace "(\d+) ", '$1","'
                $cell
            }            
            else {
              $cell = '","' + $cell
              $cell  
            }           
        }
        elseif($cells.IndexOf($_) -eq 3 )
        {
            if($cell -match "(\d+) ")
            {
                $cell = $cell -replace "(\d+) ", '$1","'
                $cell
            }            
            else {
              $cell = $cell + '","' 
              $cell  
            }           
        }
        elseif($cells.IndexOf($_) -eq 5 -or $cells.IndexOf($_) -eq 6 -or $cells.IndexOf($_) -eq 9)
        {
        $cell = $cell -replace " ", '","'
        if($cell -match "^\s*$")
            {
            $cell = $cell + '","'    
            }
        if(-not ($cell -match '\"\,\"'))
        {
            $cell = $cell + '","'
        }
        $cell
        }
        else {
            $cell
        }
    }
    $csvWriter.WriteLine('"' + ($cellValues -join '","') + '"')
}
}
# Close the CSV file
$csvWriter.Close()

Write-Host "Table data has been written to $csvFilePath"
