#requires -version 6.2
<#
    .SYNOPSIS
        This command will generate a Word document containing the information about all the Azure Sentinel
        Solutions.  
    .DESCRIPTION
       This command will generate a Word document containing the information about all the Azure Sentinel
        Solutions.  
    .PARAMETER FileName
        Enter the file name to use.  Defaults to "MicrosoftSentinelSolutions.docx"  ".docx" will be appended to all filenames if needed
    .NOTES
        AUTHOR: Gary Bushey
        LASTEDIT: 16 March 2023
    .EXAMPLE
        Export-AzSentinelSolutionsToWord 
        In this example you will get the file named "MicrosoftSentinelSolutionsReport.docx" generated containing all the solution information
    .EXAMPLE
        Export-AzSentinelSolutionsToWord -fileName "test"
        In this example you will get the file named "test.docx" generated containing all the solution information
   
#>

[CmdletBinding()]
param (

    [string]$FileName = "MicrosoftSentinelSolutionsReport.docx"
)

Function Export-AzSentinelSolutionsToWord($fileName) {
    try {
        #Setup Word
        $word = New-Object -ComObject word.application
        $word.Visible = $false      #We don't want to see word while this is running
        $doc = $word.documents.add()
        $selection = $word.Selection        #This is where we will add all the text and formatting information

        #Title Page
        $selection.Style = "Title"
        $selection.ParagraphFormat.Alignment = 1  #Center
        $selection.TypeText("Microsoft Sentinel Solutions Documentation")   #Add the text
        $selection.TypeParagraph()                                          #create a new paragraph
        $selection.TypeParagraph()
        $selection.Style = "Normal"
        $selection.ParagraphFormat.Alignment = 1  #Center
        $text = Get-Date
        $selection.TypeText("Created: " + $text)
        $selection.TypeParagraph()
        $selection.InsertBreak(7)   #page break

        #Tables of Contents.  Will need to update when everything else is added
        $range = $selection.Range
        #Note that you need to reference tableSofcontents (Note the "S" in tables, rather than tableofcontents)
        #Not going to say how long it took me to get this right ;)
        $toc = $doc.TablesOfContents.add($range, $true, 1, 1)
        $selection.TypeParagraph()
        $selection.InsertBreak(7)   #page break
    

        #Load the list of all solutions
        $url = "https://catalogapi.azure.com/offers?api-version=2018-08-01-beta&%24filter=categoryIds%2Fany%28cat%3A+cat+eq+%27AzureSentinelSolution%27%29+or+keywords%2Fany%28key%3A+contains%28key%2C%27f1de974b-f438-4719-b423-8bf704ba2aef%27%29%29"
        $solutions = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).items

        #Just used for testing and to show how many solutions we have worked on
        $count = 1

        #Go through and get each solution alphabetically
        foreach ($solution in $solutions | Sort-Object -Property "displayName") {
           # if ($count -le 70) {
                Write-Host $count  $solution.displayName
                Add-SingleSolution $solution $selection     #Load one solution
                $selection.InsertBreak(7)   #page break
                $count = $count + 1
            <# }
            else {
                break;
            }  #>
        }
        $toc.Update()   #Update the Tables of contents
        $outputPath = Join-Path $PWD.Path $fileName
        $doc.SaveAs($outputPath)    #Save the document
        $doc.Close()                #Close the document
        $word.Quit()                #Quit word
        #NOTE:  If for some reason the Powershell quits before it can quit work, go into task manager
        #to close manually, otherwise you can get some weird results
    }
    catch {
        $errorReturn = $_
        Write-Error $errorReturn
        $doc.Close()
        $word.Quit()
    }
}

#Work with a single solutions
Function Add-SingleSolution ($solution, $selection) {

    try {
        #We need to load the solution's template information, which is stored in a separate file.  Note that
        #some solutions have multiple plans associated with them and we will only work with the first one.
        $uri = ($solution.plans[0].artifacts | Where-Object -Property "name" -EQ "DefaultTemplate")
        if ($null -ne $uri) {
            $solutionData = Invoke-RestMethod -Method "Get" -Uri $uri.uri       #Load the solution's data

            #Output the Solution information into the Word Document
            $selection.Style = "Heading 1"
            $selection.TypeText($solution.displayName)
            $selection.TypeParagraph()
            #We are using the longSummary rather than the description since the description can contain HTML formatting that
            #I cannot determine who to translate into what word can understand
            $selection.TypeText($solution.longSummary);
            $selection.TypeParagraph()
    
            #The hardest part here was determining how each of the various elements were stored in the solutions
            #Load the dataconnectors
            $dataConnectors = $solutionData.resources | Where-Object { $_.properties.mainTemplate.resources.type -eq "Microsoft.OperationalInsights/workspaces/providers/dataConnectors" }
            #Load the workbooks
            $workbooks = $solutionData.resources | Where-Object { $_.properties.mainTemplate.resources.type -eq "Microsoft.Insights/workbooks" }
            #Load the Analytic Rules
            $ruleTemplates = $solutionData.resources | Where-Object { $_.properties.mainTemplate.resources.type -eq "Microsoft.SecurityInsights/AlertRuleTemplates" }
            #Load the Hunting Queries
            $huntingQueries = $solutionData.resources | Where-Object { $_.properties.mainTemplate.resources.type -eq "Microsoft.OperationalInsights/savedSearches" }
            #Load the Watchlists
            $watchlists = $solutionData.resources | Where-Object { $_.type -eq "Microsoft.OperationalInsights/workspaces/providers/Watchlists" }
            #Load the Playbooks
            $playBooks = $solutionData.resources | Where-Object { $_.properties.mainTemplate.resources.type -eq "Microsoft.Logic/workflows" }
            #Load the Parsers
            $parsers = $solutionData.resources | Where-Object { $_.properties.mainTemplate.resources.type -eq "Microsoft.OperationalInsights/workspaces/savedSearches" }

            #Output the summary line to the word document.  I know this is in a lot of solution's descriptions already
            #but it is in HTML and I cannot figure out how to easily translate it to something Word can understand.
            $selection.Font.Bold = $true
            $selection.TypeText("Data Connectors: ");
            $selection.Font.Bold = $false
            $selection.TypeText($dataConnectors.count);
            $selection.Font.Bold = $true
            $selection.TypeText(", Workbooks: ");
            $selection.Font.Bold = $false
            $selection.TypeText($workbooks.count);
            $selection.Font.Bold = $true
            $selection.TypeText(", Analytic Rules: ");
            $selection.Font.Bold = $false
            $selection.TypeText($ruleTemplates.count);
            $selection.Font.Bold = $true
            $selection.TypeText(", Hunting Queries: ");
            $selection.Font.Bold = $false
            $selection.TypeText($huntingQueries.count);
            $selection.Font.Bold = $true
            $selection.TypeText(", Watchlists: ");
            $selection.Font.Bold = $false
            $selection.TypeText($watchlists.count);
            $selection.Font.Bold = $true
            $selection.TypeText(", Playbooks: ");
            $selection.Font.Bold = $false
            $selection.TypeText($playBooks.count);
            $selection.TypeText(", Parsers: ");
            $selection.Font.Bold = $false
            $selection.TypeText($parsers.count);
    
            $selection.TypeParagraph()
            Add-DataConnectorsText $selection $dataConnectors
            Add-WorkbookText $selection $workbooks $solutionData.parameters
            Add-AnalyticRuleText $selection $ruleTemplates
            Add-HuntingQueryText $selection $huntingQueries
            Add-WatchlistText $selection $watchlists
            Add-PlaybooksText $selection $playBooks
            Add-ParsersText $selection $parsers
        }
    }
    catch {
        $errorReturn = $_
        Write-Error $errorReturn
    }
}

Function Add-DataConnectorsText ($selection, $dataConnectors) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Data Connectors")
    $selection.TypeParagraph()

    foreach ($dataConnector in $dataConnectors) {
        $displayName = $dataConnector.properties.mainTemplate.resources.properties.connectorUIConfig.title
        $description = $dataConnector.properties.mainTemplate.resources.properties.connectorUIConfig.descriptionMarkdown    

        $selection.Style = "Heading 3"
        $selection.Font.Underline = $true
        $selection.TypeText($displayName)
        $selection.TypeParagraph()

        $selection.TypeText($description)
        $selection.TypeParagraph()
    }
}

Function Add-WorkbookText ($selection, $workbooks, $parameters) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Workbooks")
    $selection.TypeParagraph()

    foreach ($workbook in $workbooks) {
        $nameVariable = $workbook.properties.mainTemplate.resources.properties[0].displayName.split("'")[1]
        $displayName = $parameters.$nameVariable.defaultValue
        $description = $workbook.properties.mainTemplate.resources.metadata.description
        $requiredDataTypes = $workbook.properties.mainTemplate.resources.properties[1].dependencies.criteria

        $selection.Style = "Heading 3"
        $selection.Font.Underline = $true
        $selection.TypeText($displayName)
        $selection.TypeParagraph()
        $selection.Style = "Normal"
        $selection.TypeText($description)
        $selection.TypeParagraph()
        $selection.Font.Bold = $true
        $selection.TypeText("Required Data Types")
        $selection.TypeParagraph()
        $selection.Font.Bold = $false
        foreach ($requiredDataType in $requiredDataTypes) {
            # `t == tab
            $selection.TypeText("`t" + $requiredDataType.contentId)
            $selection.TypeParagraph()
        }


    }
}

Function Add-AnalyticRuleText ($selection, $ruleTemplates) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Analytic Rules")
    $selection.TypeParagraph()

    foreach ($ruleTemplate in $ruleTemplates) {
        $displayName = $ruleTemplate.properties.mainTemplate.resources.properties[0].displayName
        $description = $ruleTemplate.properties.mainTemplate.resources.properties[0].description
        $requiredDataConnectors = $ruleTemplate.properties.mainTemplate.resources.properties[0].requiredDataConnectors

        $selection.Style = "Heading 3"
        $selection.Font.Underline = $true
        $selection.TypeText($displayName)
        $selection.TypeParagraph()
        $selection.Style = "Normal"
        $selection.TypeText($description)
        $selection.TypeParagraph()
        $selection.Font.Bold = $true
        $selection.TypeText("Required Data Connectors")
        $selection.TypeParagraph()
        $selection.Font.Bold = $false
        if ($null -eq $requiredDataConnectors) {
            # `t == tab
            $selection.TypeText("`t(none listed)")
            $selection.TypeParagraph()
        }
        else {
            foreach ($requiredDataConnector in $requiredDataConnectors) {
                # `t == tab
                $selection.TypeText("`t" + $requiredDataConnector.connectorId)
                $selection.TypeParagraph()
            }
        }
    }
}

Function Add-HuntingQueryText ($solution, $huntingQueries) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Hunting Queries")
    $selection.TypeParagraph()

    foreach ($huntingQuery in $huntingQueries) {
        $displayName = $huntingQuery.properties.mainTemplate.resources.properties.displayName
        $description = $huntingQuery.properties.mainTemplate.resources.properties.description

        $selection.Style = "Heading 3"
        $selection.Font.Underline = $true
        $selection.TypeText($displayName)
        $selection.TypeParagraph()
        $selection.Style = "Normal"
        $selection.TypeText($description)
        $selection.TypeParagraph()
    }
}

Function Add-WatchlistText ($selection, $watchlists) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Watchlists")
    $selection.TypeParagraph()

    foreach ($watchlist in $watchlists) {
        $displayName = $watchlist.properties.displayName
        $description = $watchlist.properties.description

        $selection.Style = "Heading 3"
        $selection.Font.Underline = $true
        $selection.TypeText($displayName)
        $selection.TypeParagraph()
        $selection.Style = "Normal"
        $selection.TypeText($description)
        $selection.TypeParagraph()
    }
}

Function Add-PlaybooksText ($selection, $playbooks) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Playbooks")
    $selection.TypeParagraph()

    foreach ($playbook in $playbooks) {
        $displayName = $playbook.properties.mainTemplate.metadata.title
        $description = $playbook.properties.mainTemplate.metadata.description
        $mainSteps = $playbook.properties.mainTemplate.metadata.mainSteps

        $selection.Style = "Heading 3"
        $selection.Font.Underline = $true
        $selection.TypeText($displayName)
        $selection.TypeParagraph()
        $selection.Style = "Normal"
        $selection.TypeText($description)
        $selection.TypeText($mainSteps)
        $selection.TypeParagraph()
    }
}

Function Add-ParsersText ($selection, $parsers) {
    $selection.TypeParagraph()
    $selection.Style = "Heading 2"
    $selection.Font.Size = 14
    $selection.TypeText("Parsers")
    $selection.TypeParagraph()

    if ($null -ne $parsers) {
        foreach ($parser in $parsers) {
            $displayName = $parser.properties.maintemplate.resources.properties.displayName
            $category = $parser.properties.maintemplate.resources.properties.category
            $alias = $parser.properties.maintemplate.resources.properties.functionAlias

            $selection.Style = "Heading 3"
            $selection.Font.Underline = $true
            $selection.TypeText($displayName)
            $selection.TypeParagraph()
            $selection.Style = "Normal"
            $selection.TypeText("Category: " + $category)
            $selection.TypeParagraph()
            $selection.TypeText("Function Alias: " + $alias)
            $selection.TypeParagraph()
        }
    }
}

#Execute the code
if (! $Filename.EndsWith(".docx")) {
    $FileName += ".docx"
}
Export-AzSentinelSolutionsToWord $FileName