#Receber Parâmetros
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Relative Site URL Ex: /sites/sitename or /teams/sitename", Position = 0)][ValidateNotNull()]
    [string]$RelativeUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Type Tenant name without .onmicrosoft.com", Position = 1)][ValidateNotNull()]
    [string]$TenantName,
    [Parameter(Mandatory = $true, HelpMessage = "Type e-mail SPO Admin", Position = 2)][ValidateNotNull()]
    [string]$Owner,
    [Parameter(Mandatory = $true, HelpMessage = "Type L&D Contributors e-mail M365 Group", Position = 3)][ValidateNotNull()]
    [string]$LDContributors,
    [Parameter(Mandatory = $true, HelpMessage = "Type L&D Approvers separeted with semicolon", Position = 4)][ValidateNotNull()]
    [string]$LDApprovers
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
$userEmail = $Owner
$CamlQuery = @"
<View>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='EMail' />
                <Value Type='Text'>$userEmail</Value>
            </Eq>
        </Where>
    </Query>
</View>
"@

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

    New-PnPTenantSite -Title "Descubra, compartilhe e priorize o aprendizado" -Url $SiteURL -Lcid 1033 -TimeZone 8 -Template "SITEPAGEPUBLISHING#0" -Owner $Owner -Wait -Connection $adminconn -ErrorAction Stop
    
    $currentsite = $SiteURL
    $currentSiteConn = Connect-PnPOnline $currentsite -Interactive -ReturnConnection

    Write-host "Aplicando o modelo de site..." -ForeGroundColor Yellow
    Start-Sleep -Seconds 30
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
    Add-PnPFolder -Name "Training Catalog" -Folder "$($RelativeUrl)/Viva Learning Catalog" -ErrorAction Stop -Connection $currentSiteConn

    # Adiciona a permissão do Grupo do M365 a pasta do repositório de conteúdo global
    $userprofile = Get-PnPListItem -List /_catalogs/users -Query $CamlQuery -Connection $currentSiteConn

    if($userprofile["MUILanguages"] -eq "pt-BR"){
        Set-PnPFolderPermission -List 'Viva Learning Catalog' -Identity 'Viva Learning Catalog/Training Catalog' -User $LDContributors -AddRole 'Leitura' -Connection $currentSiteConn
        Write-host "Permission granted successfully..." -ForegroundColor Yellow
    }else{
        Set-PnPFolderPermission -List 'Viva Learning Catalog' -Identity 'Viva Learning Catalog/Training Catalog' -User $LDContributors -AddRole 'Read' -Connection $currentSiteConn
        Write-host "Permission granted successfully..." -ForegroundColor Yellow
    }
    
    
    # Cria os registros de configurações
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "appL"; "configurationvalue" = "ba1cabe6-dfd2-4334-96c0-0dcdf86e18e5"} -Connection $currentSiteConn
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "templateInstanceId"; "configurationvalue" = "please insert GUID value"} -Connection $currentSiteConn
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "environment"; "configurationvalue" = "please insert GUID value"} -Connection $currentSiteConn
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "approvers"; "configurationvalue" = $LDApprovers} -Connection $currentSiteConn
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "vivalearningURL"; "configurationvalue" = "https://teams.microsoft.com/l/entity/2e3a628d-6f54-4100-9e7a-f00bc3621a85/2e3a628d-6f54-4100-9e7a-f00bc3621a85"} -Connection $currentSiteConn
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "appDeepLinkID"; "configurationvalue" = "https://teams.microsoft.com/l/entity/[APPID]/[APPID]"} -Connection $currentSiteConn
    Add-PnPListItem -List "Learning App Settings" -Values @{"configurationname" = "supportedExtensions"; "configurationvalue" = "pdf;mov;mp4;avi;m4a;ppt;pptx;doc;docx;xls;xlsx"} -Connection $currentSiteConn
    
    $objField = Get-PnPField -List "Learning App Settings" -Identity "Title" -Connection $currentSiteConn
    $objField.Required = $false
    $objField.Hidden = $true
    $objField.Update()
    Invoke-PnPQuery -Connection $currentSiteConn


    #Oculta library não utilizadas pela solução
    Set-PnPList -Identity "Documents" -Hidden $true -Connection $currentSiteConn
    Set-PnPList -Identity "Form Templates" -Hidden $true -Connection $currentSiteConn

    #Importa os Termos no Site para ser utilizado na Coluna SkillTags
    $termgroup = Get-PnPSiteCollectionTermStore -Connection $currentSiteConn | Select-Object Name 
    Import-PnPTermSet -GroupName $termgroup.Name -Path '.\termsetSkillTags.csv' -IsOpen $true -Contact $Owner -Owner $Owner -Connection $currentSiteConn
    
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





 

 



