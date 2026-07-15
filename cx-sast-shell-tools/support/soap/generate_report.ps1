param(
    [Parameter(Mandatory = $true)]
    [hashtable]$session,
    [Parameter(Mandatory = $true)]
    [String]$scan_id,
    [Parameter(Mandatory = $false)]
    [string]$report_type
)

$templatepath = "$PSScriptRoot"
Import-LocalizedData -BaseDirectory $templatepath -FileName CxReportTemplate.psd1 -BindingVariable ReportTemplate

if ($report_type) {
    $ReportTemplate.reportType = $report_type
}

$soap_path = "/cxwebinterface/Portal/CxWebService.asmx"

$soap_url = New-Object System.Uri $session.base_url, $soap_path

$soap_action = "http://Checkmarx.com/CreateScanReport"
#$headers = @{
#    SOAPAction = "http://Checkmarx.com/CreateScanReport";
#}

$xml_template = @"
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:chec="http://Checkmarx.com">
   <soapenv:Header/>
   <soapenv:Body>
      <chec:CreateScanReport>
         <chec:SessionID>{0}</chec:SessionID>
         <chec:Report>
            <chec:Type>{1}</chec:Type>
            <chec:ScanID>{2}</chec:ScanID>
            <chec:DisplayData>
               <chec:Queries>
                  <chec:All>{3}</chec:All>
                  <chec:IDs>
                     <!--Zero or more repetitions:-->
                     <chec:long>{4}</chec:long>
                  </chec:IDs>
               </chec:Queries>
               <chec:ResultsSeverity>
                  <chec:All>{5}</chec:All>
                  <chec:Critical>{6}</chec:Critical>
                  <chec:High>{7}</chec:High>
                  <chec:Medium>{8}</chec:Medium>
                  <chec:Low>{9}</chec:Low>
                  <chec:Info>{10}</chec:Info>
               </chec:ResultsSeverity>
               
               <chec:ResultsState>
                  <chec:All>{11}</chec:All>
                  
                  <chec:IDs>
                     <!--Zero or more repetitions:-->
                     <chec:long>{12}</chec:long>
                  </chec:IDs>
               </chec:ResultsState>
               
               <chec:DisplayCategories>
                  <chec:All>{13}</chec:All>
                  
                  <chec:IDs>
                     <!--Zero or more repetitions:-->
                     <chec:long>{14}</chec:long>
                  </chec:IDs>
               </chec:DisplayCategories>
               
               <chec:ResultsAssigedTo>
                  <chec:All>{15}</chec:All>
                  
                  <chec:IDs>
                     <!--Zero or more repetitions:-->
                     <chec:long>{16}</chec:long>
                  </chec:IDs>
                  
                  <chec:Usernames>
                     <!--Zero or more repetitions:-->
                     <chec:string>{17}</chec:string>
                  </chec:Usernames>
               </chec:ResultsAssigedTo>
               
               <chec:ResultsPerVulnerability>
                  <chec:All>{18}</chec:All>
                  <chec:Maximimum>{19}</chec:Maximimum>
               </chec:ResultsPerVulnerability>
               
               <chec:HeaderOptions>
                  <chec:Link2OnlineResults>{20}</chec:Link2OnlineResults>
                  <chec:Team>{21}</chec:Team>
                  <chec:CheckmarxVersion>{22}</chec:CheckmarxVersion>
                  <chec:ScanComments>{23}</chec:ScanComments>
                  <chec:ScanType>{24}</chec:ScanType>
                  <chec:SourceOrigin>{25}</chec:SourceOrigin>
                  <chec:ScanDensity>{26}</chec:ScanDensity>
               </chec:HeaderOptions>
               
               <chec:GeneralOption>
                  <chec:OnlyExecutiveSummary>{27}</chec:OnlyExecutiveSummary>
                  <chec:TableOfContents>{28}</chec:TableOfContents>
                  <chec:ExecutiveSummary>{29}</chec:ExecutiveSummary>
                  <chec:DisplayCategories>{30}</chec:DisplayCategories>
                  <chec:DisplayLanguageHashNumber>{31}</chec:DisplayLanguageHashNumber>
                  <chec:ScannedQueries>{32}</chec:ScannedQueries>
                  <chec:ScannedFiles>{33}</chec:ScannedFiles>
                  <chec:VulnerabilitiesDescription>{34}</chec:VulnerabilitiesDescription>
               </chec:GeneralOption>
               
               <chec:ResultsDisplayOption>
                  <chec:AssignedTo>{35}</chec:AssignedTo>
                  <chec:Comments>{36}</chec:Comments>
                  <chec:Link2Online>{37}</chec:Link2Online>
                  <chec:ResultDescription>{38}</chec:ResultDescription>
                  <chec:SnippetsMode>{39}</chec:SnippetsMode>
               </chec:ResultsDisplayOption>
            </chec:DisplayData>
         </chec:Report>
      </chec:CreateScanReport>
   </soapenv:Body>
</soapenv:Envelope>
"@.ToString()

$body = [String]::Format($xml_template, 
                        $session.soap_session.sessionID, 
                        $ReportTemplate.reportType,
                        $scan_id,
                        $ReportTemplate.queries.all,
                        $ReportTemplate.queries.ids,
                        $ReportTemplate.resultSeverity.all,
                        $ReportTemplate.resultSeverity.critical,
                        $ReportTemplate.resultSeverity.high,
                        $ReportTemplate.resultSeverity.medium,
                        $ReportTemplate.resultSeverity.low,
                        $ReportTemplate.resultSeverity.info,
                        $ReportTemplate.resultState.all,
                        $ReportTemplate.resultState.ids,
                        $ReportTemplate.displayCategories.all,
                        $ReportTemplate.displayCategories.ids,
                        $ReportTemplate.resultsAssignedTo.all,
                        $ReportTemplate.resultsAssignedTo.ids,
                        $ReportTemplate.resultsAssignedTo.usernames,
                        $ReportTemplate.resultsPerVuln.all,
                        $ReportTemplate.resultsPerVuln.max,
                        $ReportTemplate.headerOptions.link2online,
                        $ReportTemplate.headerOptions.team,
                        $ReportTemplate.headerOptions.version,
                        $ReportTemplate.headerOptions.scanComments,
                        $ReportTemplate.headerOptions.scanType,
                        $ReportTemplate.headerOptions.sourceOrigin,
                        $ReportTemplate.headerOptions.scanDensity,
                        $ReportTemplate.generalOptions.onlyExecutiveSummary,
                        $ReportTemplate.generalOptions.tableOfContents,
                        $ReportTemplate.generalOptions.exectuiveSummary,
                        $ReportTemplate.generalOptions.displayCategories,
                        $ReportTemplate.generalOptions.displayLanguageHash,
                        $ReportTemplate.generalOptions.scannedQueries,
                        $ReportTemplate.generalOptions.scannedFiles,
                        $ReportTemplate.generalOptions.vulnDescriptions,
                        $ReportTemplate.resultsDisplayOption.assignedTo,
                        $ReportTemplate.resultsDisplayOption.comments,
                        $ReportTemplate.resultsDisplayOption.link2online,
                        $ReportTemplate.resultsDisplayOption.resultsDescription,
                        $ReportTemplate.resultsDisplayOption.snippetsMode)


$response = &"$PSScriptRoot/soap_request.ps1" $session $body $soap_url $soap_action

   $content = New-Object System.Xml.XmlDocument
   $content.LoadXml($response.Content)

$reportID = $content.DocumentElement.SelectSingleNode("//*[local-name() = 'ID']").InnerText   

return $reportID