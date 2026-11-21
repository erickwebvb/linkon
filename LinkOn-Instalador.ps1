# =====================================================================
#  LinkOn - Instalador Automático do Microsoft Office
#  Fonte: Office Deployment Tool OFICIAL da Microsoft (officecdn.microsoft.com)
#  ONLINE (PowerShell como Admin):
#     Manual:      irm https://raw.githubusercontent.com/erickwebvb/linkon/main/LinkOn-Instalador.ps1 | iex
#     Teste:       $env:LINKON_TESTE=1; irm https://raw.githubusercontent.com/erickwebvb/linkon/main/LinkOn-Instalador.ps1 | iex
#     Automático:  $env:LINKON_AUTO=1; irm https://raw.githubusercontent.com/erickwebvb/linkon/main/LinkOn-Instalador.ps1 | iex
#  OFFLINE: Instalar-LinkOn.bat  /  Instalar-LinkOn-AUTOMATICO.bat
# =====================================================================

# Usuário do GitHub da LinkOn
$BASE_URL = 'https://raw.githubusercontent.com/erickwebvb/linkon/main'

$VERSAO = '1.1.0'
$Auto  = ($env:LINKON_AUTO -eq '1')  -or ($args -contains '-Auto')
$Teste = ($env:LINKON_TESTE -eq '1') -or ($args -contains '-Teste')
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------- Reabre como Administrador se necessário ----------
$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) {
    if ($PSCommandPath) {
        # rodando de arquivo (.bat / pendrive)
        $arg = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
        if ($Auto)  { $arg += ' -Auto' }
        if ($Teste) { $arg += ' -Teste' }
    } else {
        # rodando online (irm | iex): baixa de novo já como Administrador
        $cmd = "irm '$BASE_URL/LinkOn-Instalador.ps1' | iex"
        if ($Auto)  { $cmd = "`$env:LINKON_AUTO='1'; $cmd" }
        if ($Teste) { $cmd = "`$env:LINKON_TESTE='1'; $cmd" }
        $arg = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$cmd`""
    }
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arg
    return
}

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA' -and -not $PSCommandPath) {
    $cmd = "irm '$BASE_URL/LinkOn-Instalador.ps1' | iex"
    if ($Auto)  { $cmd = "`$env:LINKON_AUTO='1'; $cmd" }
    if ($Teste) { $cmd = "`$env:LINKON_TESTE='1'; $cmd" }
    Start-Process powershell.exe -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$cmd`""
    return
}

# Nitidez em telas com zoom (125%/150%)
try {
    Add-Type -Name Dpi -Namespace LinkOn -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();' -ErrorAction Stop
    [void][LinkOn.Dpi]::SetProcessDPIAware()
} catch {}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'

# ---------- Configuração ----------
$MARCA      = 'LinkOn'
$PASTA      = Join-Path $env:TEMP 'LinkOn'
$LOG        = Join-Path $PASTA ("linkon-log_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))
$ODT_URL    = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
$PASTA_ODT  = Join-Path $env:ProgramData 'LinkOn\odt'
$SETUP      = Join-Path $PASTA_ODT 'setup.exe'
$C2R_REG    = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

# Nome exibido => ID do produto | sufixo para Visio/Project
$PRODUTOS = [ordered]@{
    'Microsoft 365 Apps para Empresas (Enterprise)' = 'O365ProPlusRetail|'
    'Microsoft 365 Apps para Negócios (Business)'   = 'O365BusinessRetail|'
    'Microsoft 365 Family / Personal'               = 'O365HomePremRetail|'
    'Office 2024 Professional Plus'                 = 'ProPlus2024Retail|2024'
    'Office 2024 Home & Business'                   = 'HomeBusiness2024Retail|2024'
    'Office 2024 Home & Student'                    = 'HomeStudent2024Retail|2024'
    'Office 2021 Professional Plus'                 = 'ProPlus2021Retail|2021'
    'Office 2021 Professional'                      = 'Professional2021Retail|2021'
    'Office 2021 Home & Business'                   = 'HomeBusiness2021Retail|2021'
    'Office 2021 Home & Student'                    = 'HomeStudent2021Retail|2021'
    'Office 2019 Professional Plus'                 = 'ProPlus2019Retail|2019'
    'Office 2019 Home & Business'                   = 'HomeBusiness2019Retail|2019'
    'Office 2019 Home & Student'                    = 'HomeStudent2019Retail|2019'
    'Office 2016 Professional Plus'                 = 'ProPlusRetail|'
    'Office 2016 Home & Business'                   = 'HomeBusinessRetail|'
    'Office 2016 Home & Student'                    = 'HomeStudentRetail|'
}
$IDIOMAS = [ordered]@{
    'Português (Brasil)'      = 'pt-br'
    'Igual ao Windows'        = 'MatchOS'
    'Inglês (EUA)'            = 'en-us'
    'Espanhol'                = 'es-es'
    'Português (Portugal)'    = 'pt-pt'
}
# ID do ODT => nome exibido, marcado por padrão
$APPS = [ordered]@{
    'Word'       = @('Word', $true)
    'Excel'      = @('Excel', $true)
    'PowerPoint' = @('PowerPoint', $true)
    'Outlook'    = @('Outlook', $true)
    'OneNote'    = @('OneNote', $false)
    'Access'     = @('Access', $false)
    'Publisher'  = @('Publisher', $false)
    'Teams'      = @('Teams', $false)
    'OneDrive'   = @('OneDrive', $true)
    'Groove'     = @('OneDrive Business antigo', $false)
    'Lync'       = @('Skype for Business', $false)
}
$PASTA_PERFIL = Join-Path $env:ProgramData 'LinkOn'
New-Item -ItemType Directory -Force -Path $PASTA_PERFIL | Out-Null
$PERFIL = if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'LinkOn-Perfil.json'))) { Join-Path $PSScriptRoot 'LinkOn-Perfil.json' } else { Join-Path $PASTA_PERFIL 'LinkOn-Perfil.json' }
$PERFIL_ONLINE = "$BASE_URL/perfil.json"

# ---------- Propaganda LinkOn IPTV (mostrada durante a instalação) ----------
$SITE_PLANOS = 'https://planos-linkon.lovable.app'
$WHATSAPP    = 'https://wa.me/5592991404012?text=Ol%C3%A1!%20Vim%20pelo%20instalador%20LinkOn%20e%20quero%20conhecer%20os%20planos.'
# Se existir anuncios.json no seu GitHub, ele substitui esta lista (dá para mudar sem mexer no script)
$ANUNCIOS = @(
    @{ Tag = 'LINKON IPTV';      Titulo = 'Canais, filmes, séries e esportes em um só lugar'; Texto = 'Planos a partir de R$ 35/mês. Valores fixos até 2027.' }
    @{ Tag = 'BLACK FRIDAY';     Titulo = 'Plano Ultra de R$ 50,00 por R$ 44,90';               Texto = 'O plano mais completo, com os maiores descontos.'; Ate = '2026-11-30' }
    @{ Tag = 'TESTE GRÁTIS';     Titulo = 'Experimente grátis por 4 horas';                     Texto = 'Peça seu teste agora pelo WhatsApp.' }
    @{ Tag = 'SEM AMARRAS';      Titulo = 'Sem fidelidade e sem renovação automática';         Texto = 'Pague por Pix ou cartão. Acesso enviado pelo WhatsApp.' }
    @{ Tag = 'ONDE VOCÊ QUISER'; Titulo = 'Smart TV, celular, tablet, PC e TV Box';            Texto = 'Qualidade HD, Full HD e 4K quando disponível.' }
    @{ Tag = 'ATENDIMENTO';      Titulo = 'Suporte humano de verdade';                         Texto = 'WhatsApp (92) 99140-4012 — ajuda na configuração e nos aparelhos.' }
)
$PROCESSOS_OFFICE = 'WINWORD','EXCEL','POWERPNT','OUTLOOK','MSACCESS','MSPUB','ONENOTE','lync','VISIO','WINPROJ','officeclicktorun_setup'
New-Item -ItemType Directory -Force -Path $PASTA | Out-Null
New-Item -ItemType Directory -Force -Path $PASTA_ODT | Out-Null
# Só Administradores e SYSTEM podem gravar aqui: impede trocar o setup.exe depois da verificação
& icacls.exe $PASTA_ODT /inheritance:r /grant:r '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-18:(OI)(CI)F' 2>&1 | Out-Null

# ---------- Cores LinkOn ----------
$COR_FUNDO  = [Drawing.Color]::FromArgb(8, 5, 22)
$COR_PAINEL = [Drawing.Color]::FromArgb(22, 14, 42)
$COR_ROSA   = [Drawing.Color]::FromArgb(255, 20, 120)
$COR_TEXTO  = [Drawing.Color]::FromArgb(235, 235, 245)
$FONTE      = New-Object Drawing.Font('Segoe UI', 10)
$FONTE_B    = New-Object Drawing.Font('Segoe UI', 10, [Drawing.FontStyle]::Bold)

# ---------- Logo embutida ----------
$LOGO_B64 = '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'
$logoImg  = [Drawing.Image]::FromStream((New-Object IO.MemoryStream(,[Convert]::FromBase64String($LOGO_B64))))

# =====================================================================
#  Interface
# =====================================================================
function Novo-Grupo($texto, $x, $y, $w, $h) {
    $g = New-Object Windows.Forms.GroupBox
    $g.Text = $texto; $g.ForeColor = $COR_ROSA; $g.Font = $FONTE_B
    $g.SetBounds($x, $y, $w, $h); $form.Controls.Add($g); return $g
}
function Novo-Combo($pai, $itens, $x, $y, $w) {
    $c = New-Object Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'; $c.Font = $FONTE
    $c.BackColor = $COR_PAINEL; $c.ForeColor = $COR_TEXTO; $c.FlatStyle = 'Flat'
    $c.SetBounds($x, $y, $w, 28); [void]$c.Items.AddRange([object[]]$itens)
    $c.SelectedIndex = 0; $pai.Controls.Add($c); return $c
}
function Novo-Botao($texto, $x, $y, $w, $h, $cor, $tam = 10) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $texto; $b.Font = New-Object Drawing.Font('Segoe UI', $tam, [Drawing.FontStyle]::Bold)
    $b.BackColor = $cor; $b.ForeColor = [Drawing.Color]::White
    $b.FlatStyle = 'Flat'; $b.FlatAppearance.BorderSize = 0
    $b.Cursor = 'Hand'
    $b.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb([Math]::Min(255, $cor.R + 28), [Math]::Min(255, $cor.G + 28), [Math]::Min(255, $cor.B + 28))
    $b.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb([int]($cor.R * 0.8), [int]($cor.G * 0.8), [int]($cor.B * 0.8))
    $b.SetBounds($x, $y, $w, $h); return $b
}

$form = New-Object Windows.Forms.Form
$form.AutoScaleDimensions = New-Object Drawing.SizeF(96, 96)
$form.AutoScaleMode = 'Dpi'
$form.Text = "$MARCA Office Suite  v$VERSAO$(if ($Teste) { '  [MODO TESTE]' })"
$form.ClientSize = New-Object Drawing.Size(965, 712)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.BackColor = $COR_FUNDO
$form.ForeColor = $COR_TEXTO
$form.Font = $FONTE
try { $form.Icon = [Drawing.Icon]::FromHandle(([Drawing.Bitmap]$logoImg).GetHicon()) } catch {}

# ----- Coluna esquerda: logo, botões, progresso, log -----
$pic = New-Object Windows.Forms.PictureBox
$pic.Image = $logoImg; $pic.SizeMode = 'Zoom'
$pic.SetBounds(20, 10, 400, 260)
$form.Controls.Add($pic)

$painelAd = New-Object Windows.Forms.Panel
$painelAd.SetBounds(20, 10, 400, 260)
$painelAd.BackColor = [Drawing.Color]::FromArgb(20, 15, 32)
$painelAd.Visible = $false
$painelAd.Add_Paint({
    $g = $_.Graphics
    $pen = New-Object Drawing.Pen($COR_ROSA, 2)
    $g.DrawRectangle($pen, 1, 1, $painelAd.Width - 3, $painelAd.Height - 3)
    $pen.Dispose()
})
$form.Controls.Add($painelAd); $painelAd.BringToFront()

$adLogo = New-Object Windows.Forms.PictureBox
$adLogo.Image = $logoImg; $adLogo.SizeMode = 'Zoom'; $adLogo.SetBounds(290, 10, 100, 69)
$painelAd.Controls.Add($adLogo)
$adTag = New-Object Windows.Forms.Label
$adTag.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
$adTag.ForeColor = [Drawing.Color]::White; $adTag.BackColor = $COR_ROSA
$adTag.AutoSize = $true; $adTag.Padding = New-Object Windows.Forms.Padding(6, 3, 6, 3)
$adTag.Location = New-Object Drawing.Point(16, 18); $painelAd.Controls.Add($adTag)
$adTitulo = New-Object Windows.Forms.Label
$adTitulo.Font = New-Object Drawing.Font('Segoe UI', 17, [Drawing.FontStyle]::Bold)
$adTitulo.ForeColor = [Drawing.Color]::White
$adTitulo.SetBounds(14, 84, 372, 68); $painelAd.Controls.Add($adTitulo)
$adTexto = New-Object Windows.Forms.Label
$adTexto.Font = New-Object Drawing.Font('Segoe UI', 10)
$adTexto.ForeColor = [Drawing.Color]::FromArgb(200, 190, 220)
$adTexto.SetBounds(16, 154, 370, 40); $painelAd.Controls.Add($adTexto)
$adBtnPlanos = Novo-Botao 'VER PLANOS' 16 202 180 36 $COR_ROSA 10
$adBtnPlanos.Add_Click({ Start-Process $SITE_PLANOS })
$painelAd.Controls.Add($adBtnPlanos)
$adBtnZap = Novo-Botao 'WHATSAPP' 204 202 180 36 ([Drawing.Color]::FromArgb(37, 211, 102)) 10
$adBtnZap.Add_Click({ Start-Process $WHATSAPP })
$painelAd.Controls.Add($adBtnZap)
$adPontos = New-Object Windows.Forms.Label
$adPontos.Font = New-Object Drawing.Font('Segoe UI', 8)
$adPontos.ForeColor = $COR_ROSA; $adPontos.TextAlign = 'MiddleCenter'
$adPontos.SetBounds(16, 240, 370, 16); $painelAd.Controls.Add($adPontos)
foreach ($c in @($painelAd, $adTag, $adTitulo, $adTexto, $adLogo)) { $c.Cursor = 'Hand' }

$btn = Novo-Botao 'INSTALAR' 20 280 400 48 $COR_ROSA 13
$form.Controls.Add($btn)
$btnRemover = Novo-Botao 'Somente desinstalar o Office' 20 334 400 32 ([Drawing.Color]::FromArgb(70, 40, 100))
$form.Controls.Add($btnRemover)

$lblEtapa = New-Object Windows.Forms.Label
$lblEtapa.Font = New-Object Drawing.Font('Segoe UI Semibold', 9)
$lblEtapa.ForeColor = $COR_TEXTO; $lblEtapa.AutoEllipsis = $true
$lblEtapa.Text = 'Pronto para começar.'
$lblEtapa.SetBounds(20, 372, 340, 18); $form.Controls.Add($lblEtapa)
$lblPct = New-Object Windows.Forms.Label
$lblPct.Font = New-Object Drawing.Font('Segoe UI Semibold', 9)
$lblPct.ForeColor = $COR_ROSA; $lblPct.TextAlign = 'MiddleRight'; $lblPct.Text = '0%'
$lblPct.SetBounds(360, 372, 60, 18); $form.Controls.Add($lblPct)

# barra de progresso nas cores LinkOn
$barra = New-Object Windows.Forms.Panel
$barra.SetBounds(20, 392, 400, 8); $barra.BackColor = $COR_PAINEL; $barra.Tag = 0
$barra.Add_Paint({
    $w = [int]($barra.Width * ([int]$barra.Tag) / 100)
    if ($w -gt 0) {
        $pincel = New-Object Drawing.SolidBrush($COR_ROSA)
        $_.Graphics.FillRectangle($pincel, 0, 0, $w, $barra.Height); $pincel.Dispose()
    }
})
$barra | Add-Member -MemberType ScriptProperty -Name Value -Value { [int]$this.Tag } -SecondValue {
    param($v) $this.Tag = [Math]::Max(0, [Math]::Min(100, [int]$v)); $lblPct.Text = "$($this.Tag)%"; $this.Invalidate()
} -Force
$form.Controls.Add($barra)

$txt = New-Object Windows.Forms.TextBox
$txt.Multiline = $true; $txt.ReadOnly = $true; $txt.ScrollBars = 'Vertical'
$txt.BackColor = $COR_PAINEL; $txt.ForeColor = $COR_TEXTO; $txt.BorderStyle = 'FixedSingle'
$txt.Font = New-Object Drawing.Font('Consolas', 9)
$txt.SetBounds(20, 408, 400, 272)
$form.Controls.Add($txt)

# ----- Coluna direita: opções -----
$grpProd = Novo-Grupo 'Versão do Office' 440 10 500 70
$cmbProd = Novo-Combo $grpProd @($PRODUTOS.Keys) 15 28 335
$btnRec = Novo-Botao '★ Recomendado' 358 27 127 28 $COR_ROSA 9
$grpProd.Controls.Add($btnRec)

$grpCfg = Novo-Grupo 'Idioma e arquitetura' 440 88 500 70
$cmbIdioma = Novo-Combo $grpCfg @($IDIOMAS.Keys) 15 28 300
$cmbArq    = Novo-Combo $grpCfg @('64 bits (recomendado)', '32 bits') 325 28 160

$grpApps = Novo-Grupo 'Programas a instalar' 440 166 500 150
$checks = @{}; $i = 0
foreach ($id in $APPS.Keys) {
    $c = New-Object Windows.Forms.CheckBox
    $c.Text = $APPS[$id][0]; $c.Checked = $APPS[$id][1]
    $c.Font = $FONTE; $c.ForeColor = $COR_TEXTO
    $c.SetBounds(15 + (($i % 3) * 160), 26 + ([Math]::Floor($i / 3) * 29), 158, 26)
    $grpApps.Controls.Add($c); $checks[$id] = $c; $i++
}

$grpExtra = Novo-Grupo 'Extras (precisam de licença própria)' 440 324 500 64
$chkVisio   = New-Object Windows.Forms.CheckBox
$chkVisio.Text = 'Visio Professional';   $chkVisio.Font = $FONTE; $chkVisio.ForeColor = $COR_TEXTO
$chkVisio.SetBounds(15, 28, 200, 26);    $grpExtra.Controls.Add($chkVisio)
$chkProject = New-Object Windows.Forms.CheckBox
$chkProject.Text = 'Project Professional'; $chkProject.Font = $FONTE; $chkProject.ForeColor = $COR_TEXTO
$chkProject.SetBounds(230, 28, 220, 26);   $grpExtra.Controls.Add($chkProject)

$grpSeg = Novo-Grupo 'Segurança / anti-falhas' 440 396 500 92
$chkRemover = New-Object Windows.Forms.CheckBox
$chkRemover.Text = 'Remover versões antigas/conflitantes antes (recomendado)'
$chkRemover.Font = $FONTE; $chkRemover.ForeColor = $COR_TEXTO; $chkRemover.Checked = $true
$chkRemover.SetBounds(15, 26, 470, 26); $grpSeg.Controls.Add($chkRemover)
$chkAtualiz = New-Object Windows.Forms.CheckBox
$chkAtualiz.Text = 'Ativar atualizações automáticas do Office'
$chkAtualiz.Font = $FONTE; $chkAtualiz.ForeColor = $COR_TEXTO; $chkAtualiz.Checked = $true
$chkAtualiz.SetBounds(15, 54, 470, 26); $grpSeg.Controls.Add($chkAtualiz)

$grpAuto = Novo-Grupo 'Modo automático' 440 496 500 148
$chkFechar = New-Object Windows.Forms.CheckBox
$chkFechar.Text = 'Fechar o instalador sozinho ao terminar'
$chkFechar.Font = $FONTE; $chkFechar.ForeColor = $COR_TEXTO; $chkFechar.Checked = $true
$chkFechar.SetBounds(15, 26, 470, 24); $grpAuto.Controls.Add($chkFechar)
$chkReiniciar = New-Object Windows.Forms.CheckBox
$chkReiniciar.Text = 'Reiniciar o PC ao terminar (aviso de 60 segundos)'
$chkReiniciar.Font = $FONTE; $chkReiniciar.ForeColor = $COR_TEXTO; $chkReiniciar.Checked = $false
$chkReiniciar.SetBounds(15, 52, 470, 24); $grpAuto.Controls.Add($chkReiniciar)
$btnSalvar = Novo-Botao 'SALVAR ESTAS OPÇÕES COMO PERFIL AUTOMÁTICO' 15 84 470 34 ([Drawing.Color]::FromArgb(45, 28, 75))
$grpAuto.Controls.Add($btnSalvar)
$lblPerfil = New-Object Windows.Forms.Label
$lblPerfil.Font = New-Object Drawing.Font('Segoe UI', 8)
$lblPerfil.ForeColor = [Drawing.Color]::FromArgb(160, 150, 190)
$lblPerfil.SetBounds(15, 122, 470, 18); $grpAuto.Controls.Add($lblPerfil)

# ----- Abas: Office | Otimizador -----
$tabs = New-Object Windows.Forms.TabControl
$tabs.SetBounds(438, 4, 512, 680)
$tabs.DrawMode = 'OwnerDrawFixed'; $tabs.SizeMode = 'Fixed'
$tabs.ItemSize = New-Object Drawing.Size(170, 30)
$tabs.Add_DrawItem({
    param($s, $e)
    $sel = ($e.Index -eq $tabs.SelectedIndex)
    $pincel = New-Object Drawing.SolidBrush($(if ($sel) { $COR_ROSA } else { $COR_PAINEL }))
    $e.Graphics.FillRectangle($pincel, $e.Bounds); $pincel.Dispose()
    $sf = New-Object Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $e.Graphics.DrawString($tabs.TabPages[$e.Index].Text, $FONTE_B, [Drawing.Brushes]::White, [Drawing.RectangleF]$e.Bounds, $sf)
})
$tabOffice = New-Object Windows.Forms.TabPage; $tabOffice.Text = 'Office';     $tabOffice.BackColor = $COR_FUNDO
$tabOtim   = New-Object Windows.Forms.TabPage; $tabOtim.Text   = 'Otimizador'; $tabOtim.BackColor   = $COR_FUNDO
$tabs.TabPages.AddRange(@($tabOffice, $tabOtim))
$form.Controls.Add($tabs)
foreach ($g in @($grpProd, $grpCfg, $grpApps, $grpExtra, $grpSeg, $grpAuto)) {
    $form.Controls.Remove($g); $g.Left -= 435; $g.Top -= 6; $tabOffice.Controls.Add($g)
}

function Novo-GrupoEm($pai, $texto, $x, $y, $w, $h) {
    $g = New-Object Windows.Forms.GroupBox
    $g.Text = $texto; $g.ForeColor = $COR_ROSA; $g.Font = $FONTE_B
    $g.SetBounds($x, $y, $w, $h); $pai.Controls.Add($g); return $g
}
function Novo-Check($pai, $texto, $x, $y, $marcado) {
    $c = New-Object Windows.Forms.CheckBox
    $c.Text = $texto; $c.Font = $FONTE; $c.ForeColor = $COR_TEXTO; $c.Checked = $marcado
    $c.SetBounds($x, $y, 465, 25); $pai.Controls.Add($c); return $c
}

$grpDiag = Novo-GrupoEm $tabOtim 'Diagnóstico deste PC' 5 4 494 168
$lblDiag = New-Object Windows.Forms.Label
$lblDiag.Font = New-Object Drawing.Font('Consolas', 9); $lblDiag.ForeColor = $COR_TEXTO
$lblDiag.SetBounds(12, 22, 474, 140); $grpDiag.Controls.Add($lblDiag)

$grpOtim = Novo-GrupoEm $tabOtim 'Otimizações seguras (nenhum arquivo é excluído)' 5 178 494 290
$y0 = 24
$chkPonto    = Novo-Check $grpOtim 'Criar ponto de restauração antes (obrigatório)'               12 ($y0 + 0)   $true
$chkPonto.Enabled = $false
$chkEnergia  = Novo-Check $grpOtim 'Plano de energia de alto desempenho (só em PC de mesa)'        12 ($y0 + 28)  $true
$chkVisual   = Novo-Check $grpOtim 'Efeitos visuais leves (mantém fontes suaves e miniaturas)'     12 ($y0 + 56)  $true
$chkInicio   = Novo-Check $grpOtim 'Abrir programas de inicialização sem atraso'                   12 ($y0 + 84)  $true
$chkSugest   = Novo-Check $grpOtim 'Desativar sugestões e propagandas do Windows'                  12 ($y0 + 112) $true
$chkTelem    = Novo-Check $grpOtim 'Reduzir coleta de dados e anúncios personalizados'             12 ($y0 + 140) $true
$chkDisco    = Novo-Check $grpOtim 'Otimizar disco (TRIM no SSD / desfragmentar HD)'               12 ($y0 + 168) $true
$chkDNS      = Novo-Check $grpOtim 'Limpar cache de DNS (ajuda em lentidão de internet)'           12 ($y0 + 196) $true
$chkReparar  = Novo-Check $grpOtim 'Reparar arquivos do Windows - SFC + DISM (20 a 40 min)'        12 ($y0 + 224) $false

$chkOtimAuto = Novo-Check $tabOtim 'Otimizar também no modo automático (depois do Office)' 10 474 $true
$chkOtimAuto.Font = $FONTE_B

$btnOtimizar = Novo-Botao 'OTIMIZAR AGORA' 5 504 494 46 $COR_ROSA 13
$tabOtim.Controls.Add($btnOtimizar)
$btnDesfazer = Novo-Botao 'Desfazer a última otimização' 5 556 494 32 ([Drawing.Color]::FromArgb(70, 40, 100))
$tabOtim.Controls.Add($btnDesfazer)
$lblNota = New-Object Windows.Forms.Label
$lblNota.Font = New-Object Drawing.Font('Segoe UI', 8); $lblNota.ForeColor = [Drawing.Color]::FromArgb(160, 150, 190)
$lblNota.Text = "Não apaga arquivos, não remove programas e não desliga Windows Update, antivírus, pesquisa ou impressoras. Todas as mudanças ficam salvas e podem ser desfeitas no botão acima ou pelo ponto de restauração."
$lblNota.SetBounds(8, 594, 488, 44); $tabOtim.Controls.Add($lblNota)

$rod = New-Object Windows.Forms.Label
$rod.Text = "$MARCA Office Suite v$VERSAO  •  Instalação oficial Microsoft  •  Logs: $PASTA"
$rod.ForeColor = [Drawing.Color]::FromArgb(140, 130, 170)
$rod.Font = New-Object Drawing.Font('Segoe UI', 8)
$rod.SetBounds(20, 688, 920, 20)
$form.Controls.Add($rod)

$TODOS_CONTROLES = @($btnRemover, $tabs)
function Travar([bool]$travar) { foreach ($c in $TODOS_CONTROLES) { $c.Enabled = -not $travar } }

# =====================================================================
#  Funções
# =====================================================================
function Log([string]$msg, [int]$progresso = -1) {
    $linhaLog = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $msg
    Add-Content -Path $LOG -Value $linhaLog -Encoding UTF8
    $txt.AppendText($linhaLog + "`r`n")
    $lblEtapa.Text = $msg.Trim()
    if ($progresso -ge 0) { $barra.Value = [Math]::Min(100, $progresso) }
    [Windows.Forms.Application]::DoEvents()
}

function Aguardar([int]$ms) {
    $fim = (Get-Date).AddMilliseconds($ms)
    while ((Get-Date) -lt $fim) { Start-Sleep -Milliseconds 100; [Windows.Forms.Application]::DoEvents() }
}

function Executar([string]$exe, [string]$argumentos) {
    $p = Start-Process -FilePath $exe -ArgumentList $argumentos -PassThru -WindowStyle Hidden
    while (-not $p.HasExited) { Start-Sleep -Milliseconds 300; [Windows.Forms.Application]::DoEvents() }
    return $p.ExitCode
}

function Testar-Internet {
    foreach ($h in 'officecdn.microsoft.com', 'www.microsoft.com') {
        try {
            $t = New-Object Net.Sockets.TcpClient
            $ok = $t.ConnectAsync($h, 443).Wait(5000)
            $t.Close()
            if ($ok) { return $true }
        } catch {}
    }
    return $false
}

function Assinatura-OK([string]$arquivo) {
    if (-not (Test-Path $arquivo)) { return $false }
    $sig = Get-AuthenticodeSignature $arquivo
    return ($sig.Status -eq 'Valid' -and $sig.SignerCertificate.Subject -match 'O=Microsoft Corporation')
}

function Executar-ODT([string]$xml) {
    # confere a assinatura de novo imediatamente antes de executar
    if (-not (Assinatura-OK $SETUP)) { throw 'O instalador da Microsoft foi alterado ou corrompido. Operação bloqueada por segurança.' }
    return (Executar $SETUP "/configure `"$xml`"")
}

function Baixar-ODT {
    for ($i = 1; $i -le 3; $i++) {
        try {
            Log "Baixando instalador oficial da Microsoft (tentativa $i/3)..."
            if (Test-Path $SETUP) { Remove-Item $SETUP -Force }
            Invoke-WebRequest -Uri $ODT_URL -OutFile $SETUP -UseBasicParsing -TimeoutSec 120
            $sig = Get-AuthenticodeSignature $SETUP
            if ($sig.Status -eq 'Valid' -and $sig.SignerCertificate.Subject -match 'O=Microsoft Corporation') {
                Log 'Assinatura digital Microsoft verificada. OK.'
                return $true
            }
            Log "Assinatura inválida ($($sig.Status)). Tentando de novo..."
        } catch { Log "Falha no download: $($_.Exception.Message)" }
        Aguardar (3000 * $i)
    }
    return $false
}

function Fechar-Office {
    $abertos = Get-Process -Name $PROCESSOS_OFFICE -ErrorAction SilentlyContinue
    if ($abertos) {
        Log ("Fechando programas do Office abertos: " + (($abertos.Name | Sort-Object -Unique) -join ', '))
        $abertos | Stop-Process -Force -ErrorAction SilentlyContinue
        Aguardar 2000
    }
}

function Office-Instalado {
    $lista = @()
    if (Test-Path $C2R_REG) {
        $ids = (Get-ItemProperty $C2R_REG -ErrorAction SilentlyContinue).ProductReleaseIds
        if ($ids) { $lista += $ids -split ',' | ForEach-Object { "Click-to-Run: $_" } }
    }
    $chaves = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    Get-ItemProperty $chaves -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match '^Microsoft Office .*(20(07|10|13|16)|Professional|Standard|Home)' -and
                       $_.UninstallString -notmatch 'OfficeClickToRun' } |
        ForEach-Object { $lista += "MSI: $($_.DisplayName)" }
    return $lista | Sort-Object -Unique
}

function Gerar-XML($produto, $sufixo, $excluir, $idioma, $bits, $visio, $project, $atualiz) {
    $exc = ($excluir | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }) -join "`r`n"
    $extras = ''
    if ($visio)   { $extras += "    <Product ID=`"VisioPro$($sufixo)Retail`">`r`n      <Language ID=`"$idioma`" />`r`n    </Product>`r`n" }
    if ($project) { $extras += "    <Product ID=`"ProjectPro$($sufixo)Retail`">`r`n      <Language ID=`"$idioma`" />`r`n    </Product>`r`n" }
    $upd = if ($atualiz) { 'TRUE' } else { 'FALSE' }
    @"
<Configuration>
  <Add OfficeClientEdition="$bits" Channel="Current" MigrateArch="TRUE">
    <Product ID="$produto">
      <Language ID="$idioma" />
$exc
    </Product>
$extras  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Updates Enabled="$upd" />
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
}

function Verificar([string]$produto) {
    if (-not (Test-Path $C2R_REG)) { return $false }
    $cfg = Get-ItemProperty $C2R_REG -ErrorAction SilentlyContinue
    if (($cfg.ProductReleaseIds -split ',') -notcontains $produto) { return $false }
    if (-not $cfg.InstallationPath) { return $false }
    return (Test-Path (Join-Path $cfg.InstallationPath 'root\Office16'))
}

function Reinicio-Pendente {
    $chaves = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
              'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    foreach ($k in $chaves) { if (Test-Path $k) { return $true } }
    return $false
}

function Aguardar-OutraInstalacao {
    # Evita conflito com outra instalação/atualização do Office em andamento
    for ($i = 0; $i -lt 60; $i++) {
        $ocupado = Get-Process -Name 'OfficeClickToRun_Setup','OfficeC2RClient','setup' -ErrorAction SilentlyContinue |
                   Where-Object { $_.Path -and $_.Path -ne $SETUP -and $_.Path -match 'Office|ClickToRun' }
        if (-not $ocupado) { return }
        if ($i -eq 0) { Log 'Outra instalação do Office está em andamento. Aguardando terminar...' }
        Aguardar 5000
    }
    throw 'Outra instalação do Office não terminou após 5 minutos. Reinicie o PC e tente de novo.'
}

function Remover-ClickToRun {
    Log 'Removendo Office Click-to-Run (pode levar alguns minutos)...' 20
    $rem = Join-Path $PASTA_ODT 'remover.xml'
    @'
<Configuration>
  <Remove All="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
'@ | Out-File $rem -Encoding utf8
    $cod = Executar-ODT $rem
    Log "Remoção concluída (código $cod)." 35
    Aguardar 5000
}

function Carregar-Anuncios {
    try {
        $online = Invoke-RestMethod -Uri "$BASE_URL/anuncios.json" -UseBasicParsing -TimeoutSec 6
        if ($online -and @($online).Count -gt 0) {
            $script:ANUNCIOS = @($online | ForEach-Object { @{ Tag = $_.Tag; Titulo = $_.Titulo; Texto = $_.Texto; Ate = $_.Ate } })
        }
    } catch {}
    # remove promoções vencidas
    $hoje = (Get-Date).Date
    $script:ANUNCIOS = @($script:ANUNCIOS | Where-Object { -not $_.Ate -or ([datetime]$_.Ate) -ge $hoje })
}

$script:AD_I = 0
function Mostrar-Anuncio {
    if ($script:ANUNCIOS.Count -eq 0) { return }
    $a = $script:ANUNCIOS[$script:AD_I % $script:ANUNCIOS.Count]
    $adTag.Text = $a.Tag; $adTitulo.Text = $a.Titulo; $adTexto.Text = $a.Texto
    $adPontos.Text = (0..($script:ANUNCIOS.Count - 1) | ForEach-Object { if ($_ -eq ($script:AD_I % $script:ANUNCIOS.Count)) { '●' } else { '○' } }) -join '  '
    $script:AD_I++
}

$timerAd = New-Object Windows.Forms.Timer
$timerAd.Interval = 7000
$timerAd.Add_Tick({ Mostrar-Anuncio })
foreach ($c in @($painelAd, $adTag, $adTitulo, $adTexto, $adLogo)) { $c.Add_Click({ Start-Process $SITE_PLANOS }) }

function Propaganda([bool]$ligar) {
    if ($ligar -and $script:ANUNCIOS.Count -gt 0) {
        Mostrar-Anuncio; $painelAd.Visible = $true; $pic.Visible = $false; $timerAd.Start()
    } else {
        $timerAd.Stop(); $painelAd.Visible = $false; $pic.Visible = $true
    }
}

# =====================================================================
#  Análise do PC e recomendação automática
# =====================================================================
function Analisar-PC {
    $r = [ordered]@{ OS = 'Desconhecido'; Win10Mais = $true; Is64 = [Environment]::Is64BitOperatingSystem
                     RAM = 0; CPU = '?'; Notebook = $false; SSD = $false; TipoDisco = '?'; Saude = '?'
                     LivreGB = 0; Idioma = (Get-UICulture).Name; Atual = @(); Casa = $false
                     Versao = 'Microsoft 365 Apps para Empresas (Enterprise)'; IdiomaRec = 'Igual ao Windows'; Bits = 0
                     Motivo = '' }
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $cs = Get-CimInstance Win32_ComputerSystem
        $build = [int]$os.BuildNumber
        $r.OS = "$($os.Caption -replace 'Microsoft ', '') (build $build)"
        $r.Win10Mais = $build -ge 10240
        $r.Casa = $os.Caption -match 'Home|Single Language|Core'
        $r.RAM = [Math]::Round($cs.TotalPhysicalMemory / 1GB)
        $r.CPU = ((Get-CimInstance Win32_Processor | Select-Object -First 1).Name -replace '\s+', ' ').Trim()
        $r.Notebook = [bool](Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue)
    } catch {}
    try {
        $letra = $env:SystemDrive.TrimEnd(':')
        $r.LivreGB = [Math]::Round((Get-PSDrive $letra).Free / 1GB)
        $disco = Get-Partition -DriveLetter $letra | Get-Disk
        $pd = Get-PhysicalDisk | Where-Object { $_.DeviceId -eq [string]$disco.Number } | Select-Object -First 1
        $r.SSD = ($pd.MediaType -eq 'SSD') -or ($pd.BusType -eq 'NVMe')
        $r.TipoDisco = if ($r.SSD) { "SSD $($pd.BusType)" } elseif ($pd.MediaType -eq 'HDD') { 'HD' } else { "$($pd.MediaType)" }
        $r.Saude = switch ("$($pd.HealthStatus)") { 'Healthy' { 'Boa' } 'Warning' { 'ATENÇÃO' } 'Unhealthy' { 'RUIM - faça backup!' } default { '?' } }
    } catch {}
    if (Test-Path $C2R_REG) {
        $ids = (Get-ItemProperty $C2R_REG -ErrorAction SilentlyContinue).ProductReleaseIds
        if ($ids) { $r.Atual = @($ids -split ',') }
    }
    # --- versão recomendada ---
    $chavePorId = @{}
    foreach ($k in $PRODUTOS.Keys) { $chavePorId[($PRODUTOS[$k] -split '\|')[0]] = $k }
    $existente = $r.Atual | Where-Object { $chavePorId.ContainsKey($_) } | Select-Object -First 1
    if (-not $r.Win10Mais) {
        $r.Versao = 'Office 2016 Professional Plus'; $r.Motivo = 'Windows anterior ao 10: só o Office 2016 é compatível'
    } elseif ($existente) {
        $r.Versao = $chavePorId[$existente];         $r.Motivo = 'mesma versão já usada neste PC (mantém a licença)'
    } elseif ($r.Casa) {
        $r.Versao = 'Microsoft 365 Family / Personal'; $r.Motivo = 'Windows Home: versão doméstica'
    } else {
        $r.Versao = 'Microsoft 365 Apps para Empresas (Enterprise)'; $r.Motivo = 'Windows Pro/Empresarial'
    }
    $r.Bits = if ($r.Is64) { 0 } else { 1 }
    $r.IdiomaRec = switch -Wildcard ($r.Idioma) {
        'pt-BR' { 'Português (Brasil)' } 'pt-PT' { 'Português (Portugal)' } 'en-*' { 'Inglês (EUA)' }
        'es-*'  { 'Espanhol' }           default { 'Igual ao Windows' }
    }
    return [pscustomobject]$r
}

function Compativel([string]$versao) {
    if ($script:REC.Win10Mais) { return $true }
    return ($versao -like 'Office 2016*')
}

function Mostrar-Diagnostico {
    $R = $script:REC
    $atual = if ($R.Atual.Count) { $R.Atual -join ', ' } else { 'nenhum' }
    $lblDiag.Text = (@(
        "Windows : $($R.OS) $(if ($R.Is64) {'64'} else {'32'} ) bits"
        "CPU     : $($R.CPU)"
        "Memória : $($R.RAM) GB   |   Tipo: $(if ($R.Notebook) {'Notebook'} else {'PC de mesa'})"
        "Disco   : $($R.TipoDisco), $($R.LivreGB) GB livres, saúde $($R.Saude)"
        "Office  : $atual"
        "Recomendado: $($R.Versao)"
        "  ($($R.Motivo)), $(if ($R.Bits -eq 0) {'64'} else {'32'}) bits, $($R.IdiomaRec)"
    ) -join "`r`n")
    if ($R.RAM -gt 0 -and $R.RAM -lt 4) { $lblDiag.Text += "`r`nAVISO: menos de 4 GB de RAM deixa o Office lento." }
}

function Aplicar-Recomendado {
    $R = $script:REC
    if ($cmbProd.Items.Contains($R.Versao))      { $cmbProd.SelectedItem   = $R.Versao }
    if ($cmbIdioma.Items.Contains($R.IdiomaRec)) { $cmbIdioma.SelectedItem = $R.IdiomaRec }
    $cmbArq.SelectedIndex = $R.Bits
    foreach ($id in $checks.Keys) { $checks[$id].Checked = ($id -in 'Word', 'Excel', 'PowerPoint', 'Outlook', 'OneDrive') }
    $chkRemover.Checked = $true; $chkAtualiz.Checked = $true
    # otimizador: padrões conforme o hardware
    $chkEnergia.Checked = -not $R.Notebook
    $chkDisco.Checked   = $R.SSD          # no HD a desfragmentação demora; fica opcional
    $chkReparar.Checked = $false
}

# =====================================================================
#  Otimizador seguro (só ajustes reversíveis, nenhum arquivo é excluído)
# =====================================================================
$BACKUP_OTIM = Join-Path $PASTA_PERFIL 'otimizacao-backup.json'
$CDM = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
$AJUSTES = @{
    Visual = @(
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects'; N = 'VisualFXSetting'; V = 3; T = 'DWord' }
        @{ P = 'HKCU:\Control Panel\Desktop\WindowMetrics'; N = 'MinAnimate'; V = '0'; T = 'String' }
        @{ P = 'HKCU:\Control Panel\Desktop'; N = 'MenuShowDelay'; V = '100'; T = 'String' }
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; N = 'TaskbarAnimations'; V = 0; T = 'DWord' }
        @{ P = 'HKCU:\Software\Microsoft\Windows\DWM'; N = 'EnableAeroPeek'; V = 0; T = 'DWord' }
    )
    Inicio = @(
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize'; N = 'StartupDelayInMSec'; V = 0; T = 'DWord' }
    )
    Sugestoes = @(
        @{ P = $CDM; N = 'SilentInstalledAppsEnabled';       V = 0; T = 'DWord' }
        @{ P = $CDM; N = 'SystemPaneSuggestionsEnabled';     V = 0; T = 'DWord' }
        @{ P = $CDM; N = 'SoftLandingEnabled';               V = 0; T = 'DWord' }
        @{ P = $CDM; N = 'SubscribedContent-338388Enabled';  V = 0; T = 'DWord' }
        @{ P = $CDM; N = 'SubscribedContent-338389Enabled';  V = 0; T = 'DWord' }
        @{ P = $CDM; N = 'SubscribedContent-353694Enabled';  V = 0; T = 'DWord' }
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; N = 'ShowSyncProviderNotifications'; V = 0; T = 'DWord' }
    )
    Telemetria = @(
        @{ P = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; N = 'AllowTelemetry'; V = 1; T = 'DWord' }
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy'; N = 'TailoredExperiencesWithDiagnosticDataEnabled'; V = 0; T = 'DWord' }
        @{ P = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo'; N = 'Enabled'; V = 0; T = 'DWord' }
    )
}

function Aplicar-Ajustes($lista, $backup) {
    foreach ($a in $lista) {
        try {
            $antes = (Get-ItemProperty -Path $a.P -Name $a.N -ErrorAction SilentlyContinue).($a.N)
            [void]$backup.Add([pscustomobject]@{ P = $a.P; N = $a.N; T = $a.T; Antes = $antes; Existia = ($null -ne $antes) })
            if (-not (Test-Path $a.P)) { New-Item -Path $a.P -Force | Out-Null }
            New-ItemProperty -Path $a.P -Name $a.N -Value $a.V -PropertyType $a.T -Force | Out-Null
        } catch { Log "   (ignorado: $($a.N) - $($_.Exception.Message))" }
    }
}

function Criar-PontoRestauracao {
    Log 'Criando ponto de restauração do Windows (pode levar alguns minutos)...'
    try {
        $sr = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
        New-ItemProperty -Path $sr -Name 'SystemRestorePointCreationFrequency' -Value 0 -PropertyType DWord -Force | Out-Null
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
    } catch {}
    $cmd = "Checkpoint-Computer -Description 'LinkOn - antes da otimizacao' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop"
    $cod = Executar 'powershell.exe' "-NoProfile -ExecutionPolicy Bypass -Command `"$cmd`""
    return ($cod -eq 0)
}

function Iniciar-Otimizacao([bool]$continuacao = $false) {
    if (-not $continuacao) { Travar $true; $btn.Enabled = $false; Propaganda $true; $txt.Clear(); $barra.Value = 0 }
    $ok = $false
    try {
        Log "$MARCA - Otimizador seguro iniciado." 2
        if (-not (Criar-PontoRestauracao)) {
            Log 'Não foi possível criar o ponto de restauração.'
            if ($script:MODO_AUTO) { throw 'Sem ponto de restauração: otimização cancelada por segurança.' }
            $r = [Windows.Forms.MessageBox]::Show("Não foi possível criar o ponto de restauração.`n`nAs mudanças ainda podem ser desfeitas pelo botão 'Desfazer'. Continuar mesmo assim?", $MARCA, 'YesNo', 'Warning')
            if ($r -ne 'Yes') { throw 'Otimização cancelada pelo usuário.' }
        } else { Log 'Ponto de restauração criado. OK.' 15 }

        $backup = New-Object Collections.ArrayList
        $energiaAntes = $null
        if ($chkEnergia.Checked) {
            if ($script:REC.Notebook) { Log 'Notebook detectado: plano de energia mantido (preserva a bateria).' }
            else {
                $energiaAntes = ([regex]::Match((powercfg /getactivescheme), '[0-9a-fA-F-]{36}')).Value
                $alto = '8c5e7fda-e8bf-4a96-9a85-cf48e1e74ed6'
                powercfg /setactive $alto 2>$null
                if ($LASTEXITCODE -ne 0) {
                    $novo = ([regex]::Match((powercfg -duplicatescheme $alto), '[0-9a-fA-F-]{36}')).Value
                    if ($novo) { powercfg /setactive $novo 2>$null }
                }
                Log 'Plano de energia: alto desempenho.' 25
            }
        }
        if ($chkVisual.Checked) { Aplicar-Ajustes $AJUSTES.Visual $backup;     Log 'Efeitos visuais leves aplicados.' 35 }
        if ($chkInicio.Checked) { Aplicar-Ajustes $AJUSTES.Inicio $backup;     Log 'Inicialização sem atraso aplicada.' 40 }
        if ($chkSugest.Checked) { Aplicar-Ajustes $AJUSTES.Sugestoes $backup;  Log 'Sugestões e propagandas do Windows desativadas.' 45 }
        if ($chkTelem.Checked)  { Aplicar-Ajustes $AJUSTES.Telemetria $backup; Log 'Coleta de dados reduzida.' 50 }

        # salva o backup ANTES das etapas demoradas
        $dados = [pscustomobject]@{ Data = (Get-Date).ToString('dd/MM/yyyy HH:mm'); EnergiaAntes = $energiaAntes; Registro = $backup }
        $dados | ConvertTo-Json -Depth 4 | Out-File $BACKUP_OTIM -Encoding utf8
        Copy-Item $BACKUP_OTIM (Join-Path $PASTA_PERFIL ("otimizacao-backup_{0:yyyyMMdd_HHmmss}.json" -f (Get-Date))) -ErrorAction SilentlyContinue
        Log 'Backup das configurações salvo (pode ser desfeito).' 55

        if ($chkDNS.Checked) { Executar 'ipconfig.exe' '/flushdns' | Out-Null; Log 'Cache de DNS limpo.' 58 }
        if ($chkDisco.Checked) {
            $letra = $env:SystemDrive.TrimEnd(':')
            $modo = if ($script:REC.SSD) { '-ReTrim' } else { '-Defrag' }
            Log "Otimizando disco ${letra}: ($(if ($script:REC.SSD) {'TRIM do SSD'} else {'desfragmentação, pode demorar'}))..." 60
            $cod = Executar 'powershell.exe' "-NoProfile -Command `"Optimize-Volume -DriveLetter $letra $modo`""
            Log "Disco otimizado (código $cod)." 75
        }
        if ($chkReparar.Checked) {
            Log 'Verificando a imagem do Windows (DISM)... pode levar 20 minutos.' 78
            $cod = Executar 'DISM.exe' '/Online /Cleanup-Image /RestoreHealth'
            Log "DISM finalizado (código $cod)." 88
            Log 'Verificando arquivos do sistema (SFC)...'
            $cod = Executar 'sfc.exe' '/scannow'
            Log "SFC finalizado (código $cod). Detalhes em C:\Windows\Logs\CBS\CBS.log" 97
        }
        Log 'OTIMIZAÇÃO CONCLUÍDA! Alguns ajustes aparecem depois de reiniciar o PC.' 100
        $ok = $true
        if (-not $continuacao) { Avisar "Otimização concluída!`n`nReinicie o PC para aplicar tudo. Se quiser voltar como era, use 'Desfazer a última otimização'." $MARCA }
    } catch {
        Log "ERRO na otimização: $($_.Exception.Message)"
        if (-not $continuacao) { Avisar "$($_.Exception.Message)`n`nLog: $LOG" "$MARCA - Erro" 'Error' }
    } finally {
        if (-not $continuacao) { Travar $false; $btn.Enabled = $true; Propaganda $false }
    }
    return $ok
}

function Desfazer-Otimizacao {
    if (-not (Test-Path $BACKUP_OTIM)) { [Windows.Forms.MessageBox]::Show('Nenhuma otimização para desfazer neste PC.', $MARCA) | Out-Null; return }
    $b = Get-Content $BACKUP_OTIM -Raw -Encoding UTF8 | ConvertFrom-Json
    $r = [Windows.Forms.MessageBox]::Show("Desfazer a otimização feita em $($b.Data)?", $MARCA, 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    Travar $true; $btn.Enabled = $false; $txt.Clear()
    try {
        foreach ($i in @($b.Registro)) {
            try {
                if ($i.Existia) { New-ItemProperty -Path $i.P -Name $i.N -Value $i.Antes -PropertyType $i.T -Force | Out-Null }
                else { Remove-ItemProperty -Path $i.P -Name $i.N -ErrorAction SilentlyContinue }
            } catch { Log "   (não restaurado: $($i.N))" }
        }
        if ($b.EnergiaAntes) { powercfg /setactive $b.EnergiaAntes 2>$null; Log 'Plano de energia original restaurado.' }
        Rename-Item $BACKUP_OTIM ("otimizacao-desfeita_{0:yyyyMMdd_HHmmss}.json" -f (Get-Date)) -ErrorAction SilentlyContinue
        Log 'Configurações originais restauradas. Reinicie o PC.' 100
        [Windows.Forms.MessageBox]::Show('Pronto! Tudo voltou como era. Reinicie o PC.', $MARCA, 'OK', 'Information') | Out-Null
    } finally { Travar $false; $btn.Enabled = $true }
}

# =====================================================================
#  Autoteste (não instala, não altera nada no PC)
# =====================================================================
function Executar-Autoteste {
    Travar $true; $btn.Enabled = $false; $txt.Clear()
    $res = New-Object Collections.ArrayList
    function T([string]$nome, [scriptblock]$teste) {
        try { $det = & $teste; [void]$res.Add([pscustomobject]@{ Ok = $true;  Nome = $nome; Det = "$det" }); Log "[OK]    $nome  $det" }
        catch { [void]$res.Add([pscustomobject]@{ Ok = $false; Nome = $nome; Det = $_.Exception.Message }); Log "[FALHA] $nome  $($_.Exception.Message)" }
    }
    Log "AUTOTESTE $MARCA v$VERSAO - nada será instalado ou alterado." 1
    T 'PowerShell'            { if ($PSVersionTable.PSVersion.Major -lt 5) { throw "versão $($PSVersionTable.PSVersion) (precisa 5.1)" }; "v$($PSVersionTable.PSVersion)" }
    T 'Modo STA (janelas)'    { if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'não está em STA' }; 'OK' }
    T 'Administrador'         { 'OK' }
    T 'Análise do PC'         { if (-not $script:REC) { throw 'falhou' }; "$($script:REC.OS) | $($script:REC.RAM) GB | $($script:REC.TipoDisco)" }
    T 'Recomendação'          { if (-not $cmbProd.Items.Contains($script:REC.Versao)) { throw "versão inválida: $($script:REC.Versao)" }; $script:REC.Versao }
    T 'Pasta protegida'       { $acl = (Get-Acl $PASTA_ODT).Access | Where-Object { $_.IdentityReference -match 'Users|Usuários' -and $_.FileSystemRights -match 'Write|Modify|FullControl' }; if ($acl) { throw 'usuários comuns podem gravar' }; $PASTA_ODT }
    T 'Internet Microsoft'    { if (-not (Testar-Internet)) { throw 'sem conexão com officecdn.microsoft.com' }; 'OK' } ; $barra.Value = 30
    T 'Download + assinatura' { if (-not (Baixar-ODT)) { throw 'download ou assinatura falhou' }; (Get-AuthenticodeSignature $SETUP).SignerCertificate.Subject.Split(',')[0] } ; $barra.Value = 55
    T 'XML de todas versões'  {
        $n = 0
        foreach ($k in $PRODUTOS.Keys) {
            $id, $suf = $PRODUTOS[$k] -split '\|'
            foreach ($bits in '64', '32') {
                [xml]$x = Gerar-XML $id $suf @('Access') 'pt-br' $bits $true $true $true
                if ($x.Configuration.Add.Product[0].ID -ne $id) { throw "XML errado para $k" }
                $n++
            }
        }
        "$n configurações válidas"
    }
    T 'Compatibilidade'       { $inc = @($PRODUTOS.Keys | Where-Object { -not (Compativel $_) }); "$($PRODUTOS.Count - $inc.Count) versões compatíveis com este Windows" }
    T 'Perfil'                { if ($script:temPerfil) { $lblPerfil.Text } else { 'nenhum perfil (usará o recomendado)' } }
    T 'Anúncios'              { if ($script:ANUNCIOS.Count -eq 0) { throw 'nenhum anúncio ativo' }; "$($script:ANUNCIOS.Count) anúncios ativos" }
    T 'Ponto de restauração'  { $st = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore' -ErrorAction SilentlyContinue).RPSessionInterval; 'serviço disponível' }
    T 'Plano de energia'      { $g = ([regex]::Match((powercfg /getactivescheme), '[0-9a-fA-F-]{36}')).Value; if (-not $g) { throw 'powercfg não respondeu' }; $g }
    T 'Otimização de disco'   { if (-not (Get-Command Optimize-Volume -ErrorAction SilentlyContinue)) { throw 'Optimize-Volume indisponível' }; 'disponível' }
    T 'Ajustes de registro'   { foreach ($l in $AJUSTES.Values) { foreach ($a in $l) { if (-not ($a.P -and $a.N -and $a.T)) { throw "ajuste incompleto: $($a.N)" } } }; "$(($AJUSTES.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum) ajustes reversíveis" }
    T 'Propaganda (painel)'   { Propaganda $true; Aguardar 1500; Propaganda $false; 'exibida e ocultada' }
    $barra.Value = 100
    $falhas = @($res | Where-Object { -not $_.Ok })
    $rel = Join-Path ([Environment]::GetFolderPath('Desktop')) ("LinkOn-Autoteste_{0:yyyyMMdd_HHmm}.txt" -f (Get-Date))
    $res | ForEach-Object { "{0,-7} {1,-24} {2}" -f $(if ($_.Ok) { '[OK]' } else { '[FALHA]' }), $_.Nome, $_.Det } | Out-File $rel -Encoding utf8
    Log "Relatório salvo na Área de Trabalho: $rel"
    Travar $false; $btn.Enabled = $true
    $titulo = if ($falhas.Count) { "$MARCA - Autoteste: $($falhas.Count) falha(s)" } else { "$MARCA - Autoteste aprovado" }
    $msg = if ($falhas.Count) { "Falhas:`n" + (($falhas | ForEach-Object { "- $($_.Nome): $($_.Det)" }) -join "`n") } else { "Todos os $($res.Count) testes passaram. Pode usar nos clientes." }
    [Windows.Forms.MessageBox]::Show("$msg`n`nRelatório: $rel", $titulo, 'OK', $(if ($falhas.Count) { 'Warning' } else { 'Information' })) | Out-Null
}

function Avisar([string]$msg, [string]$titulo, [string]$icone = 'Information') {
    # No modo automático, sucesso não abre janela (só log). Erros sempre aparecem.
    if ($script:MODO_AUTO -and $icone -ne 'Error') { Log "[aviso] $msg"; return }
    [Windows.Forms.MessageBox]::Show($msg, $titulo, 'OK', $icone) | Out-Null
}

function Salvar-Perfil {
    $perfil = [ordered]@{
        Versao     = [string]$cmbProd.SelectedItem
        Idioma     = [string]$cmbIdioma.SelectedItem
        Bits       = $cmbArq.SelectedIndex
        Programas  = @($checks.Keys | Where-Object { $checks[$_].Checked })
        Visio      = $chkVisio.Checked
        Project    = $chkProject.Checked
        Remover    = $chkRemover.Checked
        Atualizar  = $chkAtualiz.Checked
        Fechar     = $chkFechar.Checked
        Reiniciar  = $chkReiniciar.Checked
        OtimizarAuto = $chkOtimAuto.Checked
        SalvoEm    = (Get-Date).ToString('dd/MM/yyyy HH:mm')
    }
    $perfil | ConvertTo-Json | Out-File $PERFIL -Encoding utf8
}

function Carregar-Perfil {
    try {
        if (Test-Path $PERFIL) {
            $p = Get-Content $PERFIL -Raw -Encoding UTF8 | ConvertFrom-Json; $origem = 'deste PC'
        } else {
            $p = Invoke-RestMethod -Uri $PERFIL_ONLINE -UseBasicParsing -TimeoutSec 6; $origem = 'online'
        }
        if (-not $p.Versao) { throw 'perfil vazio' }
        Aplicar-Recomendado   # base: recomendado para este PC
        if ($p.Versao -ne 'AUTO' -and $cmbProd.Items.Contains($p.Versao))   { $cmbProd.SelectedItem   = $p.Versao }
        if ($p.Idioma -ne 'AUTO' -and $cmbIdioma.Items.Contains($p.Idioma)) { $cmbIdioma.SelectedItem = $p.Idioma }
        if ("$($p.Bits)" -in '0', '1') { $cmbArq.SelectedIndex = [int]$p.Bits }
        if ($null -ne $p.OtimizarAuto) { $chkOtimAuto.Checked = [bool]$p.OtimizarAuto }
        foreach ($id in $checks.Keys) { $checks[$id].Checked = (@($p.Programas) -contains $id) }
        $chkVisio.Checked = [bool]$p.Visio;     $chkProject.Checked   = [bool]$p.Project
        $chkRemover.Checked = [bool]$p.Remover; $chkAtualiz.Checked   = [bool]$p.Atualizar
        $chkFechar.Checked = [bool]$p.Fechar;   $chkReiniciar.Checked = [bool]$p.Reiniciar
        $lblPerfil.Text = "Perfil $origem carregado (salvo em $($p.SalvoEm))."
        return $true
    } catch {
        $lblPerfil.Text = 'Nenhum perfil encontrado (nem neste PC, nem online).'
        return $false
    }
}

$btnSalvar.Add_Click({
    try {
        Salvar-Perfil
        Carregar-Perfil | Out-Null
        Set-Clipboard -Value (Get-Content $PERFIL -Raw)
        [Windows.Forms.MessageBox]::Show("Perfil salvo neste PC e COPIADO.`n`nPara usar em TODOS os clientes: abra o arquivo perfil.json no seu GitHub, clique no lápis (editar), apague tudo, cole (Ctrl+V) e clique em Commit changes.", $MARCA, 'OK', 'Information') | Out-Null
    } catch {
        [Windows.Forms.MessageBox]::Show("Não foi possível salvar o perfil nesta pasta: $($_.Exception.Message)", $MARCA, 'OK', 'Error') | Out-Null
    }
})

# =====================================================================
#  Instalação
# =====================================================================
function Iniciar-Instalacao {
    Travar $true; $btn.Enabled = $false
    Propaganda $true
    $txt.Clear(); $barra.Value = 0
    $sucesso = $false
    try {
        $nomeProd = [string]$cmbProd.SelectedItem
        if (-not (Compativel $nomeProd)) {
            if ($script:MODO_AUTO) {
                Log "$nomeProd não é compatível com este Windows. Usando o recomendado: $($script:REC.Versao)"
                $cmbProd.SelectedItem = $script:REC.Versao; $nomeProd = $script:REC.Versao
            } else {
                throw "$nomeProd não é compatível com $($script:REC.OS).`nRecomendado para este PC: $($script:REC.Versao)"
            }
        }
        $produto, $sufixo = $PRODUTOS[$nomeProd] -split '\|'
        $idioma   = $IDIOMAS[[string]$cmbIdioma.SelectedItem]
        $bits     = if ($cmbArq.SelectedIndex -eq 0) { '64' } else { '32' }
        $excluir  = @($checks.Keys | Where-Object { -not $checks[$_].Checked })
        if (@($checks.Keys | Where-Object { $checks[$_].Checked }).Count -eq 0) { throw 'Marque pelo menos um programa para instalar.' }
        if (($chkVisio.Checked -or $chkProject.Checked) -and $produto -like 'O365Home*') {
            throw 'Visio/Project não são compatíveis com o Microsoft 365 Family/Personal. Desmarque os extras ou escolha outra versão.'
        }
        Log "$MARCA - iniciando: $nomeProd | $idioma | $bits bits" 2
        Log ("Programas: " + (($checks.Keys | Where-Object { $checks[$_].Checked }) -join ', '))

        # 1. Verificações
        if ($bits -eq '64' -and -not [Environment]::Is64BitOperatingSystem) { throw 'Windows 32 bits detectado. Selecione "32 bits".' }
        $disco = $env:SystemDrive.TrimEnd(':')
        $livreGB = [Math]::Round((Get-PSDrive $disco).Free / 1GB, 1)
        Log "Espaço livre em ${disco}: $livreGB GB" 5
        if ($livreGB -lt 10) { throw "Espaço insuficiente. Libere pelo menos 10 GB no disco ${disco}:." }
        if (Reinicio-Pendente) { throw 'O Windows está aguardando reinicialização (atualização pendente). Reinicie o PC e rode o instalador de novo.' }
        Aguardar-OutraInstalacao
        if (-not (Testar-Internet)) { throw 'Sem conexão com os servidores da Microsoft. Verifique a internet.' }
        Log 'Internet OK.' 8

        # 2. Download do ODT
        if (-not (Baixar-ODT)) { throw 'Não foi possível baixar o instalador oficial.' }
        $barra.Value = 15

        # 3. Conflitos
        $existentes = @(Office-Instalado)
        if ($existentes.Count -gt 0) {
            Log 'Office encontrado no computador:'
            $existentes | ForEach-Object { Log "   - $_" }
        } else { Log 'Nenhum Office anterior encontrado.' }
        Fechar-Office
        if ($chkRemover.Checked -and ($existentes -match 'Click-to-Run')) { Remover-ClickToRun }
        if ($existentes -match 'MSI') { Log 'Versões MSI antigas serão removidas automaticamente durante a instalação.' }

        # 4. Instalação (com 1 nova tentativa automática)
        $xmlPath = Join-Path $PASTA_ODT 'instalar.xml'
        Gerar-XML $produto $sufixo $excluir $idioma $bits $chkVisio.Checked $chkProject.Checked $chkAtualiz.Checked |
            Out-File $xmlPath -Encoding utf8
        Copy-Item $xmlPath (Join-Path $PASTA ("instalar_{0:HHmmss}.xml" -f (Get-Date))) -ErrorAction SilentlyContinue
        for ($t = 1; $t -le 2 -and -not $sucesso; $t++) {
            Log "Instalando (tentativa $t/2). Pode levar de 10 a 40 minutos, não desligue o PC..." 40
            if (-not (Assinatura-OK $SETUP)) { throw 'O instalador da Microsoft foi alterado ou corrompido. Operação bloqueada por segurança.' }
            $p = Start-Process -FilePath $SETUP -ArgumentList "/configure `"$xmlPath`"" -PassThru -WindowStyle Hidden
            $v = 40
            while (-not $p.HasExited) {
                Start-Sleep -Milliseconds 500; [Windows.Forms.Application]::DoEvents()
                if ($v -lt 95 -and (Get-Random -Maximum 40) -eq 0) { $v++; $barra.Value = $v }
            }
            Log "Instalador finalizou (código $($p.ExitCode))."
            $sucesso = ($p.ExitCode -eq 0) -and (Verificar $produto)
            if ($p.ExitCode -ne 0) { Log "Código diferente de 0 = falha do instalador Microsoft. Logs do Office ficam em $env:TEMP" }
            if (-not $sucesso -and $t -lt 2) {
                Log 'Instalação não confirmada. Limpando e tentando novamente...'
                Fechar-Office; Remover-ClickToRun
            }
        }

        if ($sucesso) {
            Log 'INSTALAÇÃO CONCLUÍDA COM SUCESSO!' 100
            Avisar "Office instalado com sucesso!`n`nAbra o Word e entre com a conta Microsoft que possui a licença.`n`nConheça também a LinkOn IPTV: $SITE_PLANOS" $MARCA
        } else {
            throw "A instalação não pôde ser confirmada. Envie o log ao suporte $MARCA."
        }
    } catch {
        Log "ERRO: $($_.Exception.Message)"
        Log "Log salvo em: $LOG"
        Avisar "$($_.Exception.Message)`n`nLog: $LOG" "$MARCA - Erro" 'Error'
    } finally {
        Travar $false; $btn.Enabled = $true
        Propaganda $false
        $btn.Text = if ($sucesso) { 'CONCLUÍDO - INSTALAR OUTRO' } else { 'INSTALAR' }
    }

    # Otimização automática depois do Office
    if ($sucesso -and $script:MODO_AUTO -and $chkOtimAuto.Checked) {
        Travar $true; $btn.Enabled = $false; Propaganda $true
        Iniciar-Otimizacao $true | Out-Null
        Travar $false; $btn.Enabled = $true; Propaganda $false
    }

    # Ações finais (manual ou automático)
    if ($sucesso -and $chkReiniciar.Checked) {
        Log 'O PC será reiniciado em 60 segundos. Para cancelar: Win+R > shutdown /a'
        Start-Process shutdown.exe -ArgumentList '/r /t 60 /c "LinkOn: Office instalado. O PC vai reiniciar em 60 segundos."' -WindowStyle Hidden
    }
    if ($chkFechar.Checked -and $sucesso) {
        Log 'Fechando o instalador em 10 segundos...'
        Aguardar 10000
        $script:FECHAR_LIVRE = $true
        $form.Close()
    }
}

$btn.Add_Click({
    if ($script:CONTAGEM) {           # cancela a contagem do modo automático
        $timerAuto.Stop(); $script:CONTAGEM = $false; $script:MODO_AUTO = $false
        Travar $false; $btn.Text = 'INSTALAR'
        Log 'Modo automático cancelado. Ajuste as opções e clique em INSTALAR.'
        return
    }
    Iniciar-Instalacao
})

$btnRemover.Add_Click({
    $r = [Windows.Forms.MessageBox]::Show('Isso vai desinstalar TODAS as versões Click-to-Run do Office deste computador. Continuar?', $MARCA, 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }
    Travar $true; $txt.Clear(); $barra.Value = 0
    Propaganda $true
    try {
        if (-not (Testar-Internet)) { throw 'Sem internet para baixar a ferramenta oficial.' }
        if (-not (Baixar-ODT)) { throw 'Não foi possível baixar a ferramenta oficial.' }
        $existentes = @(Office-Instalado)
        if ($existentes.Count -eq 0) { Log 'Nenhum Office encontrado.' 100; return }
        $existentes | ForEach-Object { Log "   - $_" }
        Fechar-Office; Remover-ClickToRun
        $resta = @(Office-Instalado)
        if ($resta -match 'Click-to-Run') { throw 'Ainda há Office instalado. Reinicie o PC e tente novamente.' }
        if ($resta -match 'MSI') { Log 'Versões MSI antigas: remova em Configurações > Aplicativos ou instale um Office novo (remove automático).' }
        Log 'DESINSTALAÇÃO CONCLUÍDA.' 100
    } catch {
        Log "ERRO: $($_.Exception.Message)"
        [Windows.Forms.MessageBox]::Show($_.Exception.Message, "$MARCA - Erro", 'OK', 'Error') | Out-Null
    } finally { Travar $false; Propaganda $false }
})

$btnRec.Add_Click({ Aplicar-Recomendado; Log "Opções recomendadas aplicadas: $($script:REC.Versao) ($($script:REC.Motivo))." })
$btnOtimizar.Add_Click({ Iniciar-Otimizacao | Out-Null })
$btnDesfazer.Add_Click({ Desfazer-Otimizacao })

$form.Add_FormClosing({
    if (-not $script:FECHAR_LIVRE -and -not $script:CONTAGEM -and -not $btnRemover.Enabled) {
        $r = [Windows.Forms.MessageBox]::Show('A instalação está em andamento. Deseja mesmo fechar?', $MARCA, 'YesNo', 'Warning')
        if ($r -ne 'Yes') { $_.Cancel = $true }
    }
})

# =====================================================================
#  Início: carrega perfil e, no modo automático, faz tudo sozinho
# =====================================================================
$script:MODO_AUTO = $false; $script:CONTAGEM = $false; $script:FECHAR_LIVRE = $false
$script:REC = $null
$script:temPerfil = $false

$timerAuto = New-Object Windows.Forms.Timer
$timerAuto.Interval = 1000
$script:SEGUNDOS = 10
$timerAuto.Add_Tick({
    $script:SEGUNDOS--
    $btn.Text = "INSTALAÇÃO AUTOMÁTICA EM $($script:SEGUNDOS)s  (clique para cancelar)"
    if ($script:SEGUNDOS -le 0) {
        $timerAuto.Stop(); $script:CONTAGEM = $false
        Iniciar-Instalacao
    }
})

$form.Add_Shown({
    Travar $true; $btn.Enabled = $false
    $lblEtapa.Text = 'Analisando este PC...'; [Windows.Forms.Application]::DoEvents()
    $script:REC = Analisar-PC
    Mostrar-Diagnostico
    Aplicar-Recomendado
    $script:temPerfil = Carregar-Perfil
    Carregar-Anuncios
    Travar $false; $btn.Enabled = $true
    if ($Teste) { Executar-Autoteste; return }
    Log "Análise: $($script:REC.OS), $($script:REC.RAM) GB RAM, $($script:REC.TipoDisco)"
    Log "Recomendado: $($script:REC.Versao) - $($script:REC.Motivo)"
    if ($Auto) {
        if (-not $script:temPerfil) {
            Log 'Modo automático: nenhum perfil salvo. Escolha as opções, clique em "SALVAR ... PERFIL AUTOMÁTICO" e depois INSTALAR.'
            return
        }
        $script:MODO_AUTO = $true; $script:CONTAGEM = $true
        Travar $true
        Log "Modo automático LinkOn: $($cmbProd.SelectedItem)"
        Log 'A instalação começa sozinha em 10 segundos. Clique no botão rosa para cancelar.'
        $btn.Text = "INSTALAÇÃO AUTOMÁTICA EM 10s  (clique para cancelar)"
        $timerAuto.Start()
    }
})

[void]$form.ShowDialog()
