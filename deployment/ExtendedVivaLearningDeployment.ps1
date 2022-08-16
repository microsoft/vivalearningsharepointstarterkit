#Receber Parâmetros
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Relative Site URL Ex: /sites/sitename or /teams/sitename", Position = 0)][ValidateNotNull()]
    [string]$RelativeUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Type Tenant name without .onmicrosoft.com", Position = 1)][ValidateNotNull()]
    [string]$TenantName,
    [Parameter(Mandatory = $true, HelpMessage = "Type e-mail SPO Admin", Position = 2)][ValidateNotNull()]
    [string]$Owner,
    [Parameter(Mandatory = $true, HelpMessage = "Type L&D Contributors e-mail M365 Group", Position = 3)][ValidateNotNull()]
    [string]$LDContributors
)
#Exemplos:
<#
$RelativeUrl = "/sites/devtest4"
$Owner = "admin@M365x66999889.onmicrosoft.com"
$TenantName = "M365x66999889"
$LDContributors = "contentrepositoryaccess@M365x66999889.onmicrosoft.com"
#>
#Variáveis globais:
$AdminCenterURL = "https://$($TenantName)-admin.sharepoint.com"
$SiteURL = "https://$($tenantname).sharepoint.com" + $RelativeUrl
$FilePnPSiteTemplate = ".\templateVivaLearningExtendedSolutionV1_1.pnp"
$FilesPath = ".\Thumbnails"
$ServerRelativePath = "$($RelativeUrl)/vivalearningthumbnails"


#Conexão com o Admin Center do SharePoint Online
If(![string]::IsNullOrWhiteSpace($TenantName)){
    
    $adminconn = Connect-PnPOnline $AdminCenterURL -Interactive -ReturnConnection  -ErrorAction Stop  

}Else{
    
    Write-host "Paramêtro TenantName vazio, por favor preencha o parametro com o valor apropriado. Lembre-se de preencher sem espaço e sem caracteres especiais." -ForeGroundColor Red
    exit

}

If(![string]::IsNullOrWhiteSpace($SiteURL) -Or ![string]::IsNullOrWhiteSpace($Owner) -Or ![string]::IsNullOrWhiteSpace($LDContributors)){
    
    try {
    
    #Criação de Site - Certifique-se de antes executar o campo SiteUrl não contenha espaço ou caracteres especiais
    Write-host "Criando o site $($SiteURL) ..." -ForeGroundColor Yellow

    New-PnPTenantSite -Title "Descubra, compartilhe e priorize o aprendizado" -Url $SiteURL -Lcid 1046 -TimeZone 8 -Template "SITEPAGEPUBLISHING#0" -Owner $Owner -Wait -Connection $adminconn -ErrorAction Stop
    
    $currentsite = $SiteURL
    $currentSiteConn = Connect-PnPOnline $currentsite -Interactive -ReturnConnection

    Write-host "Aplicando o modelo de site..." -ForeGroundColor Yellow
    Start-Sleep -Seconds 60
    Invoke-PnPSiteTemplate -Path $FilePnPSiteTemplate -Verbose -Connection $currentSiteConn -ErrorAction Stop

    #Obtem todos os thumbnails na folder espefíficada
    $Files = Get-ChildItem -Path $FilesPath -Force -Recurse -ErrorAction Stop

    #Upload em massa das imagens na Library Viva Learning Thumbnails
    
    ForEach ($File in $Files)
    {
        Write-host "Uploading $($File.Directory)\$($File.Name)" -ForegroundColor Yellow
  
        #Upload o arquivo e preenche o campo Title
        Add-PnPFile -Path "$($File.Directory)\$($File.Name)" -Folder $ServerRelativePath -Values @{"Title" = $($File.Name)} -Connection $currentSiteConn -ErrorAction Stop
    }
    #Cria a pasta do repositório de conteúdo global
    Add-PnPFolder -Name "Training Catalog" -Folder "$($RelativeUrl)/Viva Learning Catalog" -ErrorAction Stop

    # Adiciona a permissão do Grupo do M365 a pasta do repositório de conteúdo global
    Set-PnPFolderPermission -List 'Viva Learning Catalog' -Identity 'Viva Learning Catalog/Training Catalog' -User $LDContributors -AddRole 'Read'
    
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "appL"; "configurationvalue" = "please insert GUID value"}
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "templateInstanceId"; "configurationvalue" = "please insert GUID value"}
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "environment"; "configurationvalue" = "please insert GUID value"}
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "approvers"; "configurationvalue" = "please insert emails separated with semicolon"}
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "vivalearningURL"; "configurationvalue" = "https://teams.microsoft.com/l/entity/2e3a628d-6f54-4100-9e7a-f00bc3621a85/2e3a628d-6f54-4100-9e7a-f00bc3621a85"}
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "appDeepLinkID"; "configurationvalue" = "https://teams.microsoft.com/l/entity/[APPID]/[APPID]"}


    Write-host "Criação do site criado com sucesso!!" -ForeGroundColor Green
    Write-host "Utilize o site criado para configurar no Viva Learning: $($SiteURL)" -ForeGroundColor Green
    
  }
  catch [System.Net.WebException], [System.IO.IOException] {
    $message = $_
    Write-host "Unable to apply template to $($SiteURL)" -ForeGroundColor Red
    Write-host $message -ForeGroundColor Red

  }

  Finally {

        $ErrorActionPreference = "Stop"

    }

}Else{
    
    Write-host "Por Favor preencha os parâmetros obrigatórios antes de executar este script" -ForeGroundColor Red
    exit

}





 

 



