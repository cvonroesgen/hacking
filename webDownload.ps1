# Define the URL of the web page you want to download
$url = "https://webpro.patriotproperties.com/carlisle/SearchResults.asp?page="

# Define the path to save the downloaded content
$outputFilePath = ".\patriot.html"



# Save the content to a file
$response.Content | Out-File -FilePath $outputFilePath

Write-Host "Web page downloaded and saved to $outputFilePath"

for ($i = 1; $i -lt 223; $i++) {
    Write-Host "Page $i"
    # Use Invoke-WebRequest to download the web page content
$response = Invoke-WebRequest -Uri ($url + $i)
# Save the content to a file
$response.Content | Out-File -FilePath ("patriot"+ $i + ".html")
Start-Sleep -Seconds 10
}

