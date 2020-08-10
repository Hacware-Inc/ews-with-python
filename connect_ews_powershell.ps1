
function Connect-EWS {
    param(
        [parameter(Mandatory = $True)]
        [PSCredential]
        $Credential,
        [parameter(Mandatory = $True)]
        [string]
        $Domain,
        [parameter(Mandatory = $True)]
        [string]
        $AutoDiscoverUrl

    )

    # First import the DLL, else none of the EWS classes will be available

    Try{
        Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
    }
    Catch{
        Write-Error "Error importing C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
        Write-Error $_.Exception
        Break
    }

    # Set the domain in networkcredential - don't know if this is the way to do it but it works.
    $Credential.GetNetworkCredential().Domain = $Domain

    # Create ExchangeCredential object from regular [PSCredential]
    $ExchangeCredential = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Credential.Username, $Credential.GetNetworkCredential().Password, $Credential.GetNetworkCredential().Domain)

    # Create the ExchangeService object
    $ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
    $ExchangeService.UseDefaultCredentials = $true

    # Add the ExchangeCredential to the ExchangeService
    $ExchangeService.Credentials = $ExchangeCredential

    # {$true} is needed if it's office 365 with a redirect on the autodiscovery. I think there's a lot more secure ways to check if
    # a redirection is correct.
    $ExchangeService.AutodiscoverUrl($AutoDiscoverUrl, {$true})

    # Return the object
    return $ExchangeService
}
