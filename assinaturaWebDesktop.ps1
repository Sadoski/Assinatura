# ---------------------------------------
# ???? Função: Obter Informações do Usuário Logado
# ---------------------------------------
function Get-UserInfo {
    $UserName = $env:USERNAME
    $Domain = $env:USERDOMAIN
    try {

        # Obt??m as informações do Active Directory sem módulo adicional
        $Searcher = New-Object DirectoryServices.DirectorySearcher
        $Searcher.Filter = "(&(objectClass=user)(sAMAccountName=$UserName))"
        $User = $Searcher.FindOne()

        if ($User -ne $null) {
            return @{
                Nome          = $User.Properties["displayname"]
                Email         = $User.Properties["mail"]
                Cargo         = $User.Properties["title"]
                Departamento  = $User.Properties["department"]
                Departamentos = if ($User.Properties["departamentoComplementar"]) { $User.Properties["departamentoComplementar"] -join " e " } else { "" }
                Empresa       = $User.Properties["company"]
                EnderecoEmp   = $User.Properties["physicalDeliveryOfficeName"]
                TelefoneEmp   = $User.Properties["homephone"]
                Ramal         = if ($User.Properties["telephonenumber"]) { $User.Properties["telephonenumber"] } else { "" }
                Celular       = if ($User.Properties["mobile"]) { $User.Properties["mobile"] } else { "" }
                Instagram     = $User.Properties["url"]
                Site          = $User.Properties["wwwhomepage"]
                # O campo abaixo é customizado no AD para inclusão de mais informações multivalorador
                Adicionais    = if ($User.Properties["camposAssinatura"]) { $User.Properties["camposAssinatura"] } else { "" }
                # O campo abaixo é customizado no AD para quando houver a necessidade de formato ingles (Tradução basica de alguns campos neste script)
                Ingles        = if ($User.Properties["assinaturaIngles"]) { $User.Properties["assinaturaIngles"] } else { "" }
                # O campo abaixo é customizado no AD para quando houver a necessidade de nomes alternativos
                NomeAlt       = if ($User.Properties["nomeAlternativo"]) { $User.Properties["nomeAlternativo"] } else { "" }
                # O campo abaixo é customizado no AD para quando houver a necessidade do cargo na assinatura
                CargoAtivo    = if ($User.Properties["ativarCargoAssinatura"]) { $User.Properties["ativarCargoAssinatura"] } else { "" }
                # O campo abaixo é customizado no AD para quando houver a necessidade de inclusão do banner na assinatura base64
                Banner        = if ($User.Properties["bannerAssinatura"]) { $User.Properties["bannerAssinatura"] } else { "" }
                # O campo abaixo é customizado no AD para quando houver a necessidade de inclusão do banner na assinatura URL
                BannerUrl     = if ($User.Properties["bannerURLAssinatura"]) { $User.Properties["bannerURLAssinatura"] } else { "" }
            }
        } else {
            Write-Host "$_"
            Exit 1
        }
        
    } catch {
        Write-Host "$_"
    }
}

# ---------------------------------------
# ???? Função: Obter Informações do Office Instalado
# ---------------------------------------
function Get-OfficeVersion {
    [CmdletBinding()]
    param ()

    $officeVersions = @("16.0", "15.0", "14.0")
    $basePath = "HKCU:\Software\Microsoft\Office\"

    foreach ($version in $officeVersions) {
        $path = $basePath + $version + "\Outlook"
        if (Test-Path -Path $path) {
            return $version
        }
    }

    Write-Host "$_"
    Exit 1
}


# ---------------------------------------
# ???? Função: Importar o modulo do ExchangeOnline
# ---------------------------------------
function Import-ExchangeOnlineModule {
    # Importanto o modulo de conexão com o Exchange Online
    try {
        # Pasta do modulo do Exchange Online
        $ModulePath = "\\contoso.sa\NETLOGON\Assinatura\ExchangeOnlineManagement"

        # Se o módulo ainda não foi baixado, baixa ele para a pasta temporária
        if (!(Test-Path $ModulePath)) {
            Write-Host "$_"
            Exit 1
        }

        # Copiando o modulo localmente
        Copy-Item -Path $ModulePath -Destination "$env:TEMP\ExchangeOnlineManagement" -Recurse -Force

        # Desbloqueando os arquivos
        Get-ChildItem -Path "$env:TEMP\ExchangeOnlineManagement" -Recurse | Unblock-File

        # Temporariamente define a política de execução apenas para o processo atual
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        
        # Importanto o Modulo
        Import-Module "$env:TEMP\ExchangeOnlineManagement\ExchangeOnlineManagement.psd1" -Force -DisableNameChecking

        Get-Module ExchangeOnlineManagement -ListAvailable

        Return
    } catch {
        Write-Host "$_"
    }
    
}

# ---------------------------------------
# ???? Fun????o: Para conectar no Exchange Online
# ---------------------------------------
function Open-ConnectionExchangeOnline {
    
    Try {
        # ???? Configurações de Autenticação (Azure AD)
        $AppId = "Seu Client ID"
        $TenantId = "Seu Tenant ID"
        $ClientSecret = "Sua senha"

        # Obter token de autenticação do Microsoft Entra ID (Azure AD)
        $Body = @{
            grant_type    = "client_credentials"
            client_id     = $AppId
            client_secret = $ClientSecret
            scope         = "https://outlook.office365.com/.default"
        } 
        
        $TokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $Body
        $AccessToken = $TokenResponse.access_token
        

        Connect-ExchangeOnline -AppId:$AppId -AccessToken $AccessToken -Organization:"contoso.com.br" -ShowBanner:$false

        return
    } catch {
        Write-Host "$_"
        Exit 1
    }
    
}

# ---------------------------------------
# ???? Função: Configurar Assinatura no Outlook Web (OWA)
# ---------------------------------------
function Set-OutlookWebSignature {
    param (
        [string]$Email,
        [string]$HtmlSignature,
        [string]$TextSignature
    )

    # Importando o modulo de conexão com o Exchange Online
    Import-ExchangeOnlineModule
    
    # Iniciando conexão no Exchange Online
    Open-ConnectionExchangeOnline

    try {

        # Aplicando a assinatura personalizada no Outlook Web (OWA) e New Outlook
        Set-MailboxMessageConfiguration -Identity:$Email `
                                        -AutoAddSignature:$True `
                                        -AutoAddSignatureOnReply:$True `
                                        -AutoAddSignatureOnMobile:$True `
                                        -UseDefaultSignatureOnMobile:$True `
                                        -DefaultFormat:html `
                                        -SignatureHTML:$HtmlSignature `
                                        -SignatureText:$TextSignature `
                                        -SignatureTextOnMobile:$TextSignature
        
    } catch {
        Write-Host "$_"
        # Desconectando do Exchange 
        Disconnect-ExchangeOnline -Confirm:$false
        Exit 1
    }
    # Desconectando do Exchange 
    Disconnect-ExchangeOnline -Confirm:$false
}

# ---------------------------------------
# ???? Função: Configurar Assinatura no Outlook Desktop
# ---------------------------------------
function Set-OutlookDesktopSignature {
    param (
        [string]$Email,
        [string]$HtmlSignature,
        [string]$TextSignature,
        [string]$RtfSignature
    )

    # Nome da assinatura
    $SignatureName = $Email #"assinatura"

    # Caminho da pasta de assinaturas do Outlook Desktop
    $SignaturePath = "$env:APPDATA\Microsoft\Signatures"

    # Criar a pasta se não existir
    If (!(Test-Path $SignaturePath)) {
        New-Item -ItemType Directory -Path $SignaturePath -Force
    }

    # Criar os arquivos de assinatura no Outlook Desktop
    $HtmlFile = "$SignaturePath\$SignatureName.htm"
    $RtfFile = "$SignaturePath\$SignatureName.rtf"
    $TxtFile = "$SignaturePath\$SignatureName.txt"

    $HtmlSignature | Out-File -Encoding UTF8 -FilePath $HtmlFile
    $RtfSignature | Out-File -Encoding UTF8 -FilePath $RtfFile
    $TextSignature | Out-File -FilePath $TxtFile

    try {
        # Verificando a versão do Office
        $VersionOffice = Get-OfficeVersion
        # Configurar assinatura padrão no Outlook Desktop
        # Caminho do Registro para definir a assinatura no Outlook
        $RegistryPathSettings = "HKCU:\Software\Microsoft\Office\$VersionOffice\Common\MailSettings"
        $RegistryPathGeneral = "HKCU:\Software\Microsoft\Office\$VersionOffice\Common\General"
        #$RegistryPathSetup = "HKCU:\Software\Microsoft\Office\$VersionOffice\Outlook\Setup"
        $RegistryPathProfile = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

        # Configurando Chave de registro para incluir Assinatura de e-mail padrão.
        # Pegando o nome do perfil
        $Profiles = Get-ChildItem -Path $RegistryPathProfile
        foreach ($Profile in $Profiles) {
            If ($Profile) {
                $ProfileName = $Profile.Name -replace ".*\\", ""

                # Construir o caminho completo da chave no registro
                $ProfileRegistryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$ProfileName\9375CFF0413111d3B88A00104B2A6676\00000002"

                # Verificar se a chave existe
                if (Test-Path  $ProfileRegistryPath) {
                    Set-ItemProperty -Path $ProfileRegistryPath -Name "New Signature" -Value $Email
                    Set-ItemProperty -Path $ProfileRegistryPath -Name "Reply-Forward Signature" -Value $Email
                } Else {
                    New-Item -Path $ProfileRegistryPath -Force
                    Set-ItemProperty -Path $ProfileRegistryPath -Name "New Signature" -Value $Email
                    Set-ItemProperty -Path $ProfileRegistryPath -Name "Reply-Forward Signature" -Value $Email
                }

            } Else {
                $ProfileRegistryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676"
                New-Item -Path $ProfileRegistryPath -Force
                Set-ItemProperty -Path $RegistryPath -Name "New Signature" -Value $Email
                Set-ItemProperty -Path $RegistryPath -Name "Reply-Forward Signature" -Value $Email
            }
        }

        # Criar chave para bloquear alteração de assinatura Outlook Classico
        # Criar a chave de congiuração da caixa de e-mail se não existir
        If (!(Test-Path $RegistryPathSettings)) {
            New-Item -Path $RegistryPathSettings -Force | Out-Null
        }

        Set-ItemProperty -Path $RegistryPathSettings -Name "NewSignature" -Value $SignatureName
        Set-ItemProperty -Path $RegistryPathSettings -Name "ReplySignature" -Value $SignatureName


        # Criar a chave de configuração da caixa de e-mail se não existir
        If (!(Test-Path $RegistryPathGeneral)) {
            New-Item -Path $RegistryPathGeneral -Force | Out-Null
        }

        Set-ItemProperty -Path $RegistryPathGeneral -Name "Signatures" -Value "signatures"

    } catch {
        Write-Host "$_"
        Exit 1
    }
    
}

# Limpar a pasta temporaria do usuário
$env:TEMP = "C:\Users\$env:USERNAME\AppData\Local\Temp"

Remove-Item -Path $env:TEMP\* -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------
# ???? Variavel
# ---------------------------------------
# Diretório do arquivo de assiatura
$signatureEmail = "\\contoso.sa\NETLOGON\Assinatura\assinaturas.html" # local da arquivo modelo da assinatura
$imageURL = "URL LOGO EMPRESA" # Url da imagem da logo da empresa
$imagemInstagram = "http://www.contoso.com.br/instagram.png" # Url da imagem do instagram
$imagemFacebook = "http://www.contoso.com.br/facebook.png" # Url da imagem do facebook
$imagemBase64Logo = "LOGO EMPRESA BASE64" # logo da empresa em base64
# logo do Instagram em base64
$imagemBase64Instagram = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAMAAABHPGVmAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyJpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6QkRGRUM0QUQzMzFFMTFGMDgyRUVEQjQ2RTJBQzg2OUEiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6QkRGRUM0QUUzMzFFMTFGMDgyRUVEQjQ2RTJBQzg2OUEiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDpCREZFQzRBQjMzMUUxMUYwODJFRURCNDZFMkFDODY5QSIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDpCREZFQzRBQzMzMUUxMUYwODJFRURCNDZFMkFDODY5QSIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PiM1wUAAAAMAUExURfm1t6RuyPlHAufY7flZAt0VM/pkAuMUNfklAfkbAfrr66ZBqvblYvc3AdEiV/mmAbOLzPMJGPl2AsIsdHo4q647m7I5lYA9su+4xPMwNYxIwvjUA/mIAc8jXM6u0pJWwdRtlPzszPLp8+oOJaxFkPNCQ/iXAfNmN7BpudGJr+YRK/NQN70xfP78/J1jyJtKvOsHGKNDsKk+ob4udu+HL9zG6vTLMvSjSf325X47r4tHwIlEvPS4N/fKyfzHtboygcslYMUqbtwZQ4JAtoVCufSNQ/SNkdUfU/rEAfm4AvrJAux7jdRJb5VcwPYFA+EaTJpLweK/1ag5jPbU2shPitocSfbd5LU2ivvw8fvxrfsBAcaay+9VYJJLxZ1DqJNMyfS9AXtBrfjZwvQ9NvfOAquBxfSwApRNxsR6ru+bpe/m8/TCAfjRkplqu5FNxffogvUGDeTG4O4FBtcdTvQMAbc0hbQ4jotKvswaV/778fmxAdYiTvV1fMAxg6BFtfoQAfDG041Xt/z2+MEve9QlYe8LIOx3A7kobpZLx5Y+qr4XTLk1icgoap1Hue2pCPKdbd8XPcwmaMMqYY9Fw55Is4M4s/gCD/C5AvQIEeQNNcknZNcLPLRSm/hpD/MWAfJ3Ou/h7OzU47UuhJhLvsk3aJdNwewOEfTu9+nc7/K8YJtGwZNOwvjz+PTKb88WT7E3eMMmZYhEucQoao9Kw5hLy/SABccyXdt+mvIOJ8UOR69ftPQzEPafAvVUDoRWrPNLBcgeX/KULHExpq8odeMlPdwVToNJs4lCwfkHB7kfZvdRAo9Kx/EUEd4vUOEDHIJDs/duAvy8B9IeRdYSQolMt6A3m49Bs/rNDJ9Qr4Y+vawxg6FGwJhLw7oeXKNKotQQS4lHu5lIyI5QvpVH0JdSyeJCVd9gfPGiLJNHxvPaIr9inPZeBfXdRNS73ckmVK0sk+skLt01RfOCI+SjvL1UfPUhBPTBi8JCf/THD/Y/A/rAA6cmZPPGAfDOEvcuAf///w5bs9AAAAEAdFJOU////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wBT9wclAAASVklEQVR42qSaeVzP6drHf5afxIgUmka2ZJR+ipipmFPWxsSvRItSNCNlC6UsTaJFnOJRiGJMJGuPLabGY2myZTqKiTkkxDOWI+ZxxnE85vHoXMt9f3/fnzjzOnMuyx9t7677uq/P57rv71fT2DRmjg+x73uMov+7YpxxdO++/G73cXMiQp7MfMsP1Lz5gfxnUQOvDsRoO2DpaicnJ1/fwpWRkZGhodHR0cXF1tbW7hD19UOGDBmJMYKiC4QbhKVl+JyQ/N+ABNj3Dvyfzp07MURQ6uqQAozQUKIQaAiFgjFQyi3DI3b/E0gve5d0x4zAzkDp1ElQAMOQyKaQQ4eaQICCnIiZ74I8e5FukeLoERcYz5S2SFm/enVBQR0v2BuQQyoIUTgX+FdefrLbWyFFJQeHWbi4OCqUgZLipIYIjDtiBEWAlGQgGzfL8ohdTSFBE28OG2ZhQZCM+HiuC0KWLoWqIGRKZORKQzLr1hFEYJRc5Jq5bbCcs+tNSMCORXmHDwMlhVOJZwhQoC579945cuTISgrOZh3GEEPxR16HUOXi5rZhg+XJFsaQorGLUhkCCwapBMpU2oaFhRHlzp2HFOtknBdxXcTI60ZLBrWxPFmkhiSPvZ1q45mVKRcsAyjxGYF+gLl169a1a9d8CwsLf8LYR/9h3KVYCGktrK+vrg6vru5iDIEQFIaUfGtjY+bpmYcUTMUjcXttbe+okCfzt7Uwjt0t3h5PukWMCw8PV9UFC1MeYYBYfRtrZmZmYyMojo6J2zOi5vdq/NciuUXI8nBL401W/kRCgi6ejY31IkpWXh4UJv2AfUDj74ld3ZZ3MewyhITvEpAeNf7+/gleCUBJzcvLKjvYY1vj742ikPAu1YatzAsGkIDhm4cP9xe55OWVlT1r/Hdi8EnLcFXHlM8kiDcwgOJPdfH03LFM9R26oHyMmRS98B/EzCbRK1mdTES5uvwdEBJ0Jnft2rUIwVzMDk5QvnrbM/tHW3r7+fn1vsUBe7n7cjQPddy9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"
# logo do Facebook em base64
$imagemBase64Facebook = "data:image/png;base64,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"

# Verificando se o arquivo de assinatura em html existe no local
If (-Not (Test-Path $signatureEmail)) {
    Write-Host "$_"
    Exit 1
} Else {
    # Se existir capturar as informa????es
    $assinaturaModeloHTML = Get-Content -Path $signatureEmail  -Raw
}

# ---------------------------------------
# ???? Executar as Funções
# ---------------------------------------
$UserInfo = Get-UserInfo

if ($UserInfo) {
    # Criar assinatura HTML formatada

    $temTelefone      = -not [string]::IsNullOrWhiteSpace($UserInfo.TelefoneEmp)
    $temRamal         = -not [string]::IsNullOrWhiteSpace($UserInfo.Ramal)
    $temCelular       = -not [string]::IsNullOrWhiteSpace($UserInfo.Celular)
    $temCargo         = -not [string]::IsNullOrWhiteSpace($UserInfo.Cargo)
    $temDepartamento  = -not [string]::IsNullOrWhiteSpace($UserInfo.Departamento)
    $temDepartamentos = -not [string]::IsNullOrWhiteSpace($UserInfo.Departamentos)
    $temEmail         = -not [string]::IsNullOrWhiteSpace($UserInfo.Email)
    $temEmpresa       = -not [string]::IsNullOrWhiteSpace($UserInfo.Empresa)
    $temEndereco      = -not [string]::IsNullOrWhiteSpace($UserInfo.EnderecoEmp)
    $temInstagram     = -not [string]::IsNullOrWhiteSpace($UserInfo.Instagram)
    $temSite          = -not [string]::IsNullOrWhiteSpace($UserInfo.Site)
    $temAdicionais    = -not [string]::IsNullOrWhiteSpace($UserInfo.Adicionais)
    $temNomeAlt       = -not [string]::IsNullOrWhiteSpace($UserInfo.NomeAlt)
    $temBanner        = -not [string]::IsNullOrWhiteSpace($UserInfo.Banner)
    $temBannerUrl     = -not [string]::IsNullOrWhiteSpace($UserInfo.BannerUrl)

    if ($temNomeAlt){
        if ($UserInfo.Nome) {
            $Nome = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif;color:#ef2a39;font-weight: bold;"">$($UserInfo.NomeAlt)</span>"
        } else {
            $Nome = ""
        }
    } else {
        if ($UserInfo.Nome) {
            $Nome = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif;color:#ef2a39;font-weight: bold;"">$($UserInfo.Nome)</span>"
        } else {
            $Nome = ""
        }
    }

    if ($UserInfo.CargoAtivo -eq "True"){
        if ($temCargo) {
            $Cargo = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: bold;"">$($UserInfo.Cargo)</span>"
        } else {
            $Cargo = ""
        }
    } else {
        $Cargo = ""
    }

    if ($temDepartamento -and $temCargo -and $UserInfo.CargoAtivo -eq "True") {
        $separacaoTitulo = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> - </span>"
    } else {
        $separacaoTitulo = ""
    }

    if ($temDepartamento -and $temCargo -and -not $temDepartamentos -and -not $UserInfo.CargoAtivo -eq "True") {
        $Departamento = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: bold;"">$($UserInfo.Departamento)</span>"
    } elseif ($temDepartamento -and $temDepartamentos -and $temCargo -and -not $UserInfo.CargoAtivo -eq "True") {
        $Departamento = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: bold;"">$($UserInfo.Departamento) e $($UserInfo.Departamentos)</span>"
    } elseif ($temDepartamento -and $temDepartamentos -and -not $temCargo -and -not $UserInfo.CargoAtivo -eq "True") {
        $Departamento = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: bold;"">$($UserInfo.Departamento) e $($UserInfo.Departamentos)</span>"
    } elseif ($temDepartamento -and -not $temCargo -and -not $temDepartamentos -and -not $UserInfo.CargoAtivo -eq "True") {
        $Departamento = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: bold;"">$($UserInfo.Departamento)</span>"
    } elseif ($temDepartamento -and $temCargo -and $UserInfo.CargoAtivo -eq "True" -and -not $temDepartamentos) {
        $Departamento = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: 500;"">$($UserInfo.Departamento)</span>"
    } elseif ($temDepartamento -and $temDepartamentos -and $temCargo -and $UserInfo.CargoAtivo -eq "True") {
        $Departamento = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: 500;"">$($UserInfo.Departamento) e $($UserInfo.Departamentos)</span>"
    } elseif ($temDepartamento  -and $temDepartamentos -and $UserInfo.CargoAtivo -eq "True" -and -not $temCargo) {
        $Departamento = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: 500;"">$($UserInfo.Departamento) e $($UserInfo.Departamentos)</span>"
    } elseif ($temDepartamento -and $UserInfo.CargoAtivo -eq "True" -and -not $temCargo -and -not $temDepartamentos) {
        $Departamento = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#7F7F7F;font-weight: 500;font-weight: 500;"">$($UserInfo.Departamento)</span>"
    } else {
        $Departamento = ""
    }

    if ($temEmail) {
        $Email = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;"">E-mail|Teams:</span><a style='text-decoration: none;' href='mailto:$($UserInfo.Email)'><span style=""font-size:9pt; font-family:Tahoma,sans-serif;color:#2b2526;font-weight: 500;""> $($UserInfo.Email)</span></a><br>"
    } else {
        $Email = ""
    }

    if ($temEmpresa) {
        $Empresa = "<br><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;"">$($UserInfo.Empresa)</span>"
    } else {
        $Empresa = ""
    }

    # Não utiliza mais para
    if ($temEndereco) {
        $Endereco = "<span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;"">$($UserInfo.EnderecoEmp)</span><br>"
    } else {
        $Endereco = "" 
    }

    if ($temInstagram) {
        $ImageInstagram = "<br><a style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight:500; text-decoration: none; "" href=""https://www.instagram.com/contoso/"" target=""_blank""><img src=""$imagemInstagram"" width=""9"" height=""9"" alt=""Instagram:"" style=""width:9px; height:9px;""></a>"
        $ImageInstagramBase64 = "<br><a style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight:500; text-decoration: none; width:9px; height:9px;"" href=""https://www.instagram.com/contoso/"" target=""_blank"" ><img src=""$imagemBase64Instagram"" width=""9"" height=""9"" style=""width:9px; height:9px; display:block; line-height:normal; vertical-align:middle;"" alt=""Instagram:""></a>"
    } else {
        $ImageInstagram = ""
    }

    if ($temInstagram) {
        $Instagram = "<span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;"">&nbsp;$($UserInfo.Instagram)</span>"
    } else {
        $Instagram = ""
    }

    if ($temInstagram) {
        $ImageFacebook = "<a style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight:500; text-decoration: none; width:9px; height:9px;"" width=""9"" width=""9"" "" href=""https://www.facebook.com/contoso/"" target=""_blank"">&nbsp;<img src=""$imagemFacebook"" width=""9"" height=""9"" style=""width:9px; height:9px;"" alt=""Facebook:""></a>"
        $ImageFacebookBase64 = "<a style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight:500; text-decoration: none; width:9px; height:9px;"" href=""https://www.facebook.com/contoso/"" target=""_blank"">&nbsp;<img src=""$imagemBase64Facebook"" width=""9"" height=""9"" style=""width:9px; height:9px; display:block; line-height:normal; vertical-align:middle;"" alt=""Facebook:""></a>"
    } else {
        $ImageFacebook = ""
    }

    if ($temSite) {
        $Site = "<br><a style=""font-size:9pt;font-family:'Tahoma',sans-serif;color: #f02a3b;font-weight: bold;text-decoration: none;"" href=""https://www.contoso.com.br"" target=""_blank"">$($UserInfo.Site)</a>"
    } else {
        $Site = ""
    }

    if ($imageURL){
        $LogoURL = "<img src=""$($imageURL)"" width=""80"" height=""80"" style=""width:80px; height:80px;"">"
    } else {
        $LogoURL = ""
    }

    if ($imagemBase64Logo) {
        $LogoBase64 = "<img src=""$($imagemBase64Logo)"" width=""80"" height=""80"" style=""width:80px; height:80px;"">"
    } else {
        $LogoBase64 = ""
    }

    if ($temAdicionais) {
        if ($UserInfo.Adicionais -is [array]) {
            $adicionaisArray = $UserInfo.Adicionais
        } else {
            $adicionaisArray = @($UserInfo.Adicionais)
        }

        # Verifica se tem adicionais
        $temAdicionais = $adicionaisArray.Count -gt 0


        $Adicionais = ($adicionaisArray | ForEach-Object {"<br><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;"">$_</span>"}) -join "`n"
    } else {
        $Adicionais = ""
    }


    if ($temBanner) {
        if ($UserInfo.Banner -is [array]) {

            $bannerb64 = @()
            foreach ($path in $UserInfo.Banner) {
                if (Test-Path -Path $path) {
                    $base64 = Get-Content $path -Raw
                    $bannerb64 += $base64
                }
            }
            $bannerArray = $bannerb64
        } else {
            if (Test-Path -Path $UserInfo.Banner) {
                $base64 = Get-Content $UserInfo.Banner -Raw
                $bannerArray += @($base64)
            }
        }

        # Verifica se tem Banner
        $temBanner = $bannerArray.Count -gt 0

        $BannerBase64 = ($bannerArray | ForEach-Object {"<br><img src=""$_"" width=""800"" height=""120"" style=""width:800px; height:120px;"">"}) -join "`n"
    } else {
        $BannerBase64 = ""
    }


    if ($temBannerUrl) {
        if ($UserInfo.BannerUrl -is [array]) {
            $bannerUrlArray = $UserInfo.BannerUrl
        } else {
            $bannerUrlArray = @($UserInfo.BannerUrl)
        }

        # Verifica se tem Banner 
        $temBannerUrl = $bannerUrlArray.Count -gt 0


        $BannerUrl = ($bannerUrlArray | ForEach-Object {"<br><img src=""$_"" width=""800"" height=""120"" style=""width:800px; height:120px;"">"}) -join "`n"
    } else {
        $BannerUrl = ""
    }


    if ($UserInfo.Ingles -eq "True"){
        if (-not $temRamal) {
            $Ramal = ""
        } elseif ($temRamal -and -not $temTelefone) {
            $Ramal = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif;color:#2b2526;font-weight: bold;"">Ext.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;""> $($UserInfo.Ramal)</span>"
        } else {
            $Ramal = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif;color:#2b2526;font-weight: bold;"">Ext.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;""> $($UserInfo.Ramal)</span>" 
        }

        if ($temRamal -and $temTelefone) {
            $separacaoRamal = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> &bull; </span>"
        } else {
            $separacaoRamal = ""
        }

        if ($temTelefone) {
            $Telefone = "<br><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: bold;"">Phone:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;""> $($UserInfo.TelefoneEmp)</span>"
        } else {
            $Telefone = ""
        }

        if ($temCelular -and $temRamal) {
            $separacaoCelular = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> &bull; </span>"
        } elseif ($temCelular -and $temTelefone) {
            $separacaoCelular = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> &bull; </span>"
        } else {
            $separacaoCelular = ""
        }

        if (-not $temCelular) {
            $Celular = ""
        } elseif ($temCelular -and -not $temTelefone -and -not $temRamal) {
            $Celular = "<br><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: bold;"">Mobile:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500; ""> $($UserInfo.Celular)</span>"
        } else {
            $Celular = "<span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: bold;"">Mobile:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500; ""> $($UserInfo.Celular)</span>"
        }

    } else {
        if (-not $temRamal) {
            $Ramal = ""
        } elseif ($temRamal -and -not $temTelefone) {
            $Ramal = "<br><span style=""font-size:9pt;font-family:Tahoma,sans-serif;color:#2b2526;font-weight: bold;"">R.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;""> $($UserInfo.Ramal)</span>"
        } else {
            $Ramal = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif;color:#2b2526;font-weight: bold;"">R.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;""> $($UserInfo.Ramal)</span>" 
        }

        if ($temRamal -and $temTelefone) {
            $separacaoRamal = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> &bull; </span>"
        } else {
            $separacaoRamal = ""
        }

        if ($temTelefone) {
            $Telefone = "<br><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: bold;"">Tel.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500;""> $($UserInfo.TelefoneEmp)</span>"
        } else {
            $Telefone = ""
        }

        if ($temCelular -and $temRamal) {
            $separacaoCelular = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> &bull; </span>"
        } elseif ($temCelular -and $temTelefone) {
            $separacaoCelular = "<span style=""font-size:9pt;font-family:Tahoma,sans-serif; color:#2b2526;font-weight: 500;""> &bull; </span>"
        } else {
            $separacaoCelular = ""
        }

        if (-not $temCelular) {
            $Celular = ""
        } elseif ($temCelular -and -not $temTelefone -and -not $temRamal) {
            $Celular = "<br><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: bold;"">Cel.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500; ""> $($UserInfo.Celular)</span>"
        } else {
            $Celular = "<span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: bold;"">Cel.:</span><span style=""font-size:9pt;font-family:'Tahoma',sans-serif; color:#2b2526;font-weight: 500; ""> $($UserInfo.Celular)</span>"
        }
        
    }

    # Criando a assinatura personalizada substituindo as variáveis no HTML
    $assinaturahtml = $assinaturaModeloHTML -replace "{{NOME}}", $Nome `
                                            -replace "{{CARGO}}", $Cargo `
                                            -replace "{{SEPARACAOTITULO}}", $separacaoTitulo `
                                            -replace "{{DEPARTAMENTO}}", $Departamento `
                                            -replace "{{EMAIL}}", $Email `
                                            -replace "{{RAMAL}}", $Ramal `
                                            -replace "{{SEPARACAORAMAL}}", $separacaoRamal `
                                            -replace "{{SEPARACAOCELULAR}}", $separacaoCelular `
                                            -replace "{{CELULAR}}", $Celular `
                                            -replace "{{EMPRESA}}", $Empresa `
                                            -replace "{{PHONE}}", $Telefone `
                                            -replace "{{INSTAGRAM}}", $Instagram `
                                            -replace "{{SITE}}", $Site `
                                            -replace "{{ADICIONAIS}}", $Adicionais
    
    # Criando a assinatura em texto Puro
    $assinaturaLines = @()
                                            
    if ($UserInfo.Nome) {$assinaturaLines += $UserInfo.Nome}

    if ($temCargo) { if ($UserInfo.CargoAtivo -eq "True") {$assinaturaLines += $UserInfo.Cargo}}

    if ($temDepartamento) {$assinaturaLines += $UserInfo.Departamento}

    if ($temEmpresa)      { $assinaturaLines += $UserInfo.Empresa}

    # if ($temEndereco)  { $assinaturaLines += $UserInfo.EnderecoEmp}

    # Combinar Tel e Ramal na mesma linha se algum existir
    if ($UserInfo.TelefoneEmp -or $UserInfo.Ramal -or $UserInfo.Celular) {
        $linhaTel = ""
        if ($temTelefone) { $linhaTel += "Tel: $($UserInfo.TelefoneEmp)" }
        if ($UserInfo.Ramal) {
            if ($linhaTel) { $linhaTel += "  " }
            $linhaTel += "R.: $($UserInfo.Ramal)"
        }
        if ($UserInfo.Celular) {
            if ($linhaTel) { $linhaTel += "  " }
            $linhaTel += "Cel.: $($UserInfo.Celular)"
        }
        $assinaturaLines += $linhaTel
    }

    if ($temEmail) {$assinaturaLines += "Email: $($UserInfo.Email)"}


    if ($temInstagram) {$assinaturaLines += "Instagram: $($UserInfo.Instagram)"} 

    if ($temInstagram) {$assinaturaLines += "Facebook: $($UserInfo.Instagram)"} 

    if ($temAdicionais) {
        if ($UserInfo.Adicionais -is [array]) {
            $adicionaisArray = $UserInfo.Adicionais
        } else {
            $adicionaisArray = @($UserInfo.Adicionais)
        }

        # Verifica se tem adicionais
        $temAdicionais = $adicionaisArray.Count -gt 0
        $assinaturaLines += ($adicionaisArray | ForEach-Object {$_}) -join "`n"} 

    if ($temSite) {$assinaturaLines += $UserInfo.Site} 
                                            
    
    # Junta as linhas com quebras de linha
    $assinaturaText = $assinaturaLines -join "`n"
   

    $assinaturahtmlimageurl = $assinaturahtml -replace "{{LOGO}}", $LogoURL `
                                              -replace "{{IMAGEINSTAGRAM}}", $ImageInstagram `
                                              -replace "{{IMAGEFACEBOOK}}", $ImageFacebook `
                                              -replace "{{BANNER}}", $BannerUrl


    # Aplicar no Outlook Web (OWA)
    Set-OutlookWebSignature -Email $UserInfo.Email `
                            -HtmlSignature  $assinaturahtmlimageurl `
                            #-TextSignature $assinaturaText
    

    $assinaturahtmlimagebase64 = $assinaturahtml -replace "{{LOGO}}", $LogoBase64 `
                                                 -replace "{{IMAGEINSTAGRAM}}", $ImageInstagramBase64 `
                                                 -replace "{{IMAGEFACEBOOK}}", $ImageFacebookBase64 `
                                                 -replace "{{BANNER}}", $BannerBase64

    

    # Aplicar no Outlook Desktop
    Set-OutlookDesktopSignature `
        -Email $UserInfo.Email `
        -HtmlSignature $assinaturahtmlimagebase64 `
        -TextSignature $assinaturaText `
        -RtfSignature $assinaturahtmlimagebase64
    
    # Add-Type -AssemblyName "System.Windows.Forms"
    # [System.Windows.Forms.MessageBox]::Show("Processo de assinatura conclu??do com sucesso!", "Conclus??o", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

}