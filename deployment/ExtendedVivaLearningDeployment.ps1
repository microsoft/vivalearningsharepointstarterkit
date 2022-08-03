#Receber Parâmetros
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Relative Site URL Ex: /sites/sitename or /teams/sitename", Position = 0)][ValidateNotNull()]
    [string]$RelativeUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Type Tenant name without .onmicrosoft.com", Position = 1)][ValidateNotNull()]
    [string]$TenantName,
    [Parameter(Mandatory = $true, HelpMessage = "Type SPO Admin e-mail", Position = 2)][ValidateNotNull()]
    [string]$Owner
)
#Exemplos:
<#$RelativeUrl = "/sites/devtest4"
$Owner = "admin@M365x66999889.onmicrosoft.com"
$TenantName = "M365x66999889"
#>
#Variáveis globais:
$AdminCenterURL = "https://$($TenantName)-admin.sharepoint.com"
$SiteURL = "https://$($tenantname).sharepoint.com" + $RelativeUrl
$FilePnPSiteTemplate = ".\templateVivaLearningExtendedSolution2.pnp"
$FilesPath = ".\Thumbnails"
$ServerRelativePath = "$($RelativeUrl)/vivalearningthumbnails"


#Conexão com o Admin Center do SharePoint Online
If(![string]::IsNullOrWhiteSpace($TenantName)){
    
    $adminconn = Connect-PnPOnline $AdminCenterURL -Interactive -ReturnConnection  -ErrorAction Stop  

}Else{
    
    Write-host "Paramêtro TenantName vazio, por favor preencha o parametro com o valor apropriado. Lembre-se de preencher sem espaço e sem caracteres especiais." -ForeGroundColor Red
    exit

}

If(![string]::IsNullOrWhiteSpace($SiteURL)){
    
    try {
    
    #Criação de Site - Certifique-se de antes executar o campo SiteUrl não contenha espaço ou caracteres especiais
    New-PnPTenantSite -Title "Descubra, compartilhe e priorize o aprendizado" -Url $SiteURL -Lcid 1046 -TimeZone 8 -Template "SITEPAGEPUBLISHING#0" -Owner $Owner -Wait -Connection $adminconn -ErrorAction Stop
    
    $currentsite = $SiteURL
    $currentSiteConn = Connect-PnPOnline $currentsite -Interactive -ReturnConnection
    
    Invoke-PnPSiteTemplate -Path $FilePnPSiteTemplate -Verbose -Connection $currentSiteConn -ErrorAction Stop

    #Obtem todos os thumbnails na folder espefíficada
    $Files = Get-ChildItem -Path $FilesPath -Force -Recurse

    #Upload em massa das imagens na Library Viva Learning Thumbnails
    
    ForEach ($File in $Files)
    {
        Write-host "Uploading $($File.Directory)\$($File.Name)"
  
        #Upload o arquivo e preenche o campo Title
        Add-PnPFile -Path "$($File.Directory)\$($File.Name)" -Folder $ServerRelativePath -Values @{"Title" = $($File.Name)} -Connection $currentSiteConn -ErrorAction Stop
    }

    Write-host "Criação do site criado com sucesso!!" -ForeGroundColor Green
    Write-host "Utilize o site criado para configurar no Viva Learning: $($SiteURL)" -ForeGroundColor Green
    
  }
  catch [System.Net.WebException], [System.IO.IOException] {
   
    Write-host "Unable to apply template to $($SiteURL)" -ForeGroundColor Red

  }

  Finally {

        $ErrorActionPreference = "Stop"

    }

}Else{
    
    Write-host "Paramêtro SiteURL vazio, por favor preencha o parametro com o valor apropriado. Lembre-se de preencher sem espaço e sem caracteres especiais." -ForeGroundColor Red
    exit

}





 

 



