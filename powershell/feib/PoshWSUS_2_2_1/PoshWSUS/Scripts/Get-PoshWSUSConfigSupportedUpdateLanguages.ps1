function Get-PoshWSUSConfigSupportedUpdateLanguages {
    
    <#
	.SYNOPSIS
		Gets the languages that WSUS supports.

	.DESCRIPTION
		Gets the languages that WSUS supports. The language codes follow the format specified in RFC1766. 
        For example, "en" for English or "pt-br" for Portuguese (Brazil). WSUS supports a subset of the 
        language codes that are specified in ISO 639 and culture codes that are specified in ISO 3166.

	.EXAMPLE
		Get-PoshWSUSConfigSupportedUpdateLanguages
        
        Description
        -----------  
        Gets the languages that WSUS supports.

	.OUTPUTS
		System.Collections.Specialized.StringCollection

	.NOTES
        Name: Get-PoshWSUSConfigSupportedUpdateLanguages
        Author: Dubinsky Evgeny
        DateCreated: 1DEC2013
        Modified: 06 Feb 2014 -- Boe Prox
            -Removed instances where set actions are occuring

	.LINK
		http://blog.itstuff.in.ua/?p=62#Get-PoshWSUSConfigSupportedUpdateLanguages

	.LINK
		http://msdn.microsoft.com/en-us/library/windows/desktop/microsoft.updateservices.administration.iupdateserverconfiguration.supportedupdatelanguages(v=vs.85).aspx


    #>

    [CmdletBinding()]
    Param
    (
    )
    if($wsus)
    {
        Write-Warning "Use Connect-PoshWSUSServer for establish connection with your Windows Update Server"
        Break
    } 
    Write-Verbose "Getting WSUS Supported Update Languages"
    $config.SupportedUpdateLanguages
}