try {

    Enable-WindowsOptionalFeature -Online -FeatureName "Printing-PrintToPDFServices-Features" -NoRestart -ErrorAction Stop
    
}
catch {
    
    Write-Error "Already installed! "
}