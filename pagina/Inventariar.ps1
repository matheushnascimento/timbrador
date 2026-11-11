# Inventariar.ps1 - Timbrador (white label)
# Varre a pasta dos documentos (subpastas com as siglas de tipo, Obsoletos e Projetos/Nome; Modelos fica de fora),
# le cada .docx e monta o lista-mestra.json. Mesmas regras do gerar_documento.py --inventariar, inclusive o
# SHA-256 das rodadas do relatorio de conducao.
# Nao instala nada: usa so o que ja vem no Windows PowerShell 5.1 ou superior.
# A organizacao (nome, politica, empresa padrao, tipos extras) vem do marca.json: -Marca, a variavel
# TIMBRADOR_MARCA ou um marca.json na mesma pasta deste script. Nenhum dado de empresa fica aqui.
#
#   .\Inventariar.ps1 -Pasta "G:\Documentos"
#   .\Inventariar.ps1 -Pasta "G:\Documentos" -Marca "G:\Controle\marca.json"
#   .\Inventariar.ps1 -Pasta "G:\Documentos" -Conferir      # so mostra o relatorio, nao grava
#   .\Inventariar.ps1 -Pasta "G:\Documentos" -Lista "G:\Controle"   # grava a lista em outro lugar

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)][string]$Pasta,
  [string]$Lista,
  [string]$Marca,
  [string[]]$Ignorar = @(),
  [switch]$Conferir
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$TIPOS_CURTOS = @("CI","COM","OFI","DEC","TER","ATA","RO","AUT","NOT","CTR","PRO","FOR","REC")
$TIPOS_LONGOS = @("MAN","POL","REG","POP","INS","PLN","PRJ","REL")
# A ficha embutida tem namespace terminado em "timbrador:ficha" (urn:timbrador:ficha e o dos documentos antigos).
$MARCA_FICHA  = "timbrador:ficha"
$RE_CODIGO    = [regex]"\b([A-Z0-9]{2,6})-([A-Z]{2,3})-([A-Z]{2,3})-(\d{3}|XXX)(?:/(\d{4}))?\s+·\s+(?:v|[Vv]ersão\s+)([0-9][0-9.]*(?:-r\d+)?)"
$RE_NOME      = [regex]"(?i)^([A-Z0-9]{2,6})-([A-Z]{2,4})-([A-Z]{2,4})-(\d{3}|XXX)(?:-(\d{4}))?_v([0-9][0-9.]*(?:-r\d+)?)_([A-Z_]+)_(.+)\.docx$"
$TARJAS       = [ordered]@{ "DOCUMENTO OFICIAL" = "OFICIAL"; "OBSOLETO" = "OBSOLETO"; "EM APROVAÇÃO" = "EM_APROVACAO"; "EM REVISÃO" = "EM_REVISAO"; "MINUTA" = "MINUTA"; "RASCUNHO" = "RASCUNHO" }
$ORDEM_ESTADO = @{ "OBSOLETO" = 0; "RASCUNHO" = 1; "MINUTA" = 2; "EM_REVISAO" = 3; "EM_APROVACAO" = 3; "OFICIAL" = 4 }

function Valor($obj, [string]$nome) {
  if ($null -eq $obj) { return $null }
  $p = $obj.PSObject.Properties[$nome]
  if ($null -eq $p) { return $null }
  if ($p.Value -is [string] -and $p.Value.Trim() -eq "") { return $null }
  return $p.Value
}

function Texto-Zip($zip, [string]$padrao) {
  $saida = New-Object System.Text.StringBuilder
  foreach ($e in $zip.Entries) {
    if ($e.FullName -match $padrao) {
      $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
      [void]$saida.Append(" ").Append($sr.ReadToEnd())
      $sr.Close()
    }
  }
  return $saida.ToString()
}

function Chave-Versao([string]$v) {
  if (-not $v) { $v = "0" }
  $partes = $v -split "-r", 2
  $nums = ($partes[0] -split "\.")
  $x = 0; $y = 0
  if ($nums.Count -ge 1 -and $nums[0] -match "^\d+$") { $x = [int]$nums[0] }
  if ($nums.Count -ge 2 -and $nums[1] -match "^\d+$") { $y = [int]$nums[1] }
  $temRodada = $partes.Count -gt 1
  $rod = 0
  if ($temRodada -and $partes[1] -match "^\d+$") { $rod = [int]$partes[1] }
  # oficial (sem -r) vale mais que qualquer rodada da mesma versao
  return @($x, $y, $(if ($temRodada) { 0 } else { 1 }), $rod)
}

function Compara-Versao($a, $b) {
  $A = Chave-Versao $a; $B = Chave-Versao $b
  for ($i = 0; $i -lt 4; $i++) { if ($A[$i] -ne $B[$i]) { return $A[$i] - $B[$i] } }
  return 0
}

function Ficha-Do-Docx([string]$caminho) {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($caminho)
  try {
    $spec = [ordered]@{}; $origem = "sem_timbrado"
    $xml = Texto-Zip $zip "^customXml/item\d+\.xml$"
    if ($xml -and $xml.Contains($MARCA_FICHA)) {
      $m = [regex]::Match($xml, "<ficha[^>]*>(.*?)</ficha>", "Singleline")
      if ($m.Success) {
        $json = $m.Groups[1].Value -replace "&quot;", '"' -replace "&lt;", "<" -replace "&gt;", ">" -replace "&amp;", "&"
        try {
          $f = $json | ConvertFrom-Json
          foreach ($p in $f.PSObject.Properties) { $spec[$p.Name] = $p.Value }
          $origem = "ficha"
        } catch { }
      }
    }
    if ($origem -ne "ficha") {
      $cab = (Texto-Zip $zip "^word/(header|footer)\d*\.xml$") -replace "<[^>]+>", " " -replace "\s+", " "
      $m = $RE_CODIGO.Match($cab)
      if ($m.Success) {
        $origem = "cabecalho"
        $spec["empresa"] = $m.Groups[1].Value; $spec["tipo"] = $m.Groups[2].Value
        $spec["area"] = $m.Groups[3].Value;    $spec["versao"] = $m.Groups[6].Value
        if ($m.Groups[4].Value -ne "XXX") { $spec["numero"] = [int]$m.Groups[4].Value }
        if ($m.Groups[5].Success) { $spec["ano"] = [int]$m.Groups[5].Value }
        foreach ($t in $TARJAS.Keys) { if ($cab.Contains($t)) { $spec["estado"] = $TARJAS[$t]; break } }
        if ($cab -cmatch "\bCONFIDENCIAL\b") { $spec["classificacao"] = "Confidencial" }
        $ma = [regex]::Match($cab, "Aprovado por (.+?) em (\d{2}/\d{2}/\d{4})")
        if ($ma.Success) { $spec["aprovado_por"] = $ma.Groups[1].Value.Trim(); $spec["data_aprovacao"] = $ma.Groups[2].Value }
      }
    }
    return @{ spec = $spec; origem = $origem }
  } finally { $zip.Dispose() }
}

function Ficha-Do-Nome([string]$nome) {
  $m = $RE_NOME.Match($nome)
  if (-not $m.Success) { return $null }
  $spec = [ordered]@{
    empresa = $m.Groups[1].Value.ToUpper(); tipo = $m.Groups[2].Value.ToUpper()
    area = $m.Groups[3].Value.ToUpper();    versao = $m.Groups[6].Value
    estado = $m.Groups[7].Value.ToUpper()
  }
  $titulo = ($m.Groups[8].Value -replace "-", " ").Trim()
  if ($titulo) { $spec["titulo"] = $titulo.Substring(0,1).ToUpper() + $titulo.Substring(1) }
  if ($m.Groups[4].Value.ToUpper() -ne "XXX") { $spec["numero"] = [int]$m.Groups[4].Value }
  if ($m.Groups[5].Success) { $spec["ano"] = [int]$m.Groups[5].Value }
  return $spec
}

function Sigla-Da-Pasta([string]$nome) {
  # aceita "POP", "Politicas (POL)", "REG - regulamentos"
  foreach ($m in [regex]::Matches($nome.ToUpper(), "[A-Z]{2,4}")) {
    if ($TIPOS_CURTOS -contains $m.Value -or $TIPOS_LONGOS -contains $m.Value) { return $m.Value }
  }
  return $null
}

function Chave-Registro($codigo, $numero, $arquivo) {
  # documento com numero e unico pelo codigo; sem numero (XXX), cada arquivo vale como uma entrada
  if ($numero) { return "$codigo" }
  return "$codigo|$arquivo"
}

function Sem-Acento([string]$t) {
  if (-not $t) { return "" }
  $d = $t.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($c in $d.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
  }
  return (($sb.ToString().ToLower()) -replace "[^a-z0-9]+", " ").Trim()
}

function Limpo($t) {
  # espaços colapsados e Unicode em NFC, como o _limpo do gerador: é sobre este texto que a impressão digital é feita
  if ($null -eq $t) { return "" }
  return (("$t").Normalize([Text.NormalizationForm]::FormC) -replace "\s+", " ").Trim()
}

function Pasta-Obsoletos([string]$rel) {
  foreach ($p in ("$rel" -split "[\\/]")) { if ($p.ToLower() -eq "obsoletos") { return $true } }
  return $false
}

function Projeto-Da-Pasta([string]$pasta) {
  # "Projetos/Novo Site/..." -> "Novo Site"
  $partes = @("$pasta" -split "[\\/]")
  for ($i = 0; $i -lt $partes.Count - 1; $i++) {
    if ((Sem-Acento $partes[$i]) -eq "projetos" -and $partes[$i + 1]) { return $partes[$i + 1] }
  }
  return ""
}

# ---------------------------------------------------------------- relatorio de conducao (REL com rodadas)
# Mesmas regras do gerador (celulas_rodada, digital_rodada): o SHA-256 de cada rodada e calculado sobre o texto
# das cinco celulas da linha no Quadro de rodadas, juntas por quebra de linha.
$EVENTOS_REL = @{ abertura = ""; rodada = ""; replanejamento = "Replanejamento"; parada = "Parada declarada";
                  retomada = "Retomada"; correcao = "Correção"; encerramento = "Encerramento" }
$SHA256 = [System.Security.Cryptography.SHA256]::Create()

function Celulas-Rodada($r) {
  $tipo = "$(Valor $r 'tipo')"; if (-not $tipo) { $tipo = "rodada" }
  if ($tipo -eq "replanejamento") { $pre = "Replanejamento: plano passa à v" + ((Limpo (Valor $r "plano_versao")) -replace "^[vV]", "") + "." }
  elseif ($tipo -eq "correcao") { $pre = "Correção da rodada " + (Limpo (Valor $r "corrige")) + ":" }
  elseif ($tipo -eq "encerramento") { $pre = "Encerramento (" + (Limpo (Valor $r "situacao")) + "), com o termo " + (Limpo (Valor $r "termo")) + "." }
  elseif ($EVENTOS_REL[$tipo]) { $pre = $EVENTOS_REL[$tipo] + ":" }
  else { $pre = "" }
  $c = Valor $r "compromisso"
  $ate = Limpo (Valor $c "ate")
  $partes = @(@((Limpo (Valor $c "oque")), (Limpo (Valor $c "quem")), $(if ($ate) { "até $ate" } else { "" })) | Where-Object { $_ })
  $comp = if ($partes.Count -gt 0) { $partes -join " · " } else { "—" }
  $av = "$(Valor $r 'avanco')"
  if ($av -eq "sim") { $avanco = "Sim" }
  elseif ($av -eq "nao") {
    $dep = Limpo (Valor $r "depende")
    $avanco = "Não: " + (Limpo (Valor $r "travou")) + $(if ($dep) { "; depende de $dep" } else { "" })
  } else { $avanco = "—" }
  return @(("" + [int](Valor $r "n")), (Limpo (Valor $r "data")), (Limpo ($pre + " " + (Limpo (Valor $r "feito")))), $comp, (Limpo $avanco))
}

function Digital-Rodada($r) {
  $bytes = [Text.Encoding]::UTF8.GetBytes(((Celulas-Rodada $r) -join "`n"))
  return (($SHA256.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join "")
}

# devolve sempre uma lista (vazia, se nao houver rodadas)
function Digitais-Rodadas($rodadas) {
  $lista = New-Object System.Collections.Generic.List[string]
  foreach ($r in @($rodadas)) { if ($null -ne $r) { $lista.Add((Digital-Rodada $r)) } }
  return ,$lista
}

function Encerramento-Rel($rodadas) {
  foreach ($r in @($rodadas)) { if ($null -ne $r -and "$(Valor $r 'tipo')" -eq "encerramento") { return $r } }
  return $null
}

function Tem-Rodadas($spec) {
  return ("$($spec['tipo'])" -eq "REL" -and $spec.Contains("rodadas") -and @($spec["rodadas"] | Where-Object { $null -ne $_ }).Count -gt 0)
}

# ---------------------------------------------------------------- marca.json
# Mesmas regras do gerador (gerar_documento.py --inventariar --marca): empresa padrao, nome da organizacao,
# codigo da politica e tipos extras vem dele.
if (-not $Marca) { $Marca = $env:TIMBRADOR_MARCA }
if (-not $Marca) { $Marca = Join-Path $PSScriptRoot "marca.json" }
if (-not (Test-Path -LiteralPath $Marca)) {
  Write-Host "marca.json nao encontrado ($Marca). Informe -Marca, defina TIMBRADOR_MARCA ou deixe o marca.json na pasta deste script."
  exit 1
}
try { $cfg = (Get-Content -LiteralPath $Marca -Raw -Encoding UTF8) | ConvertFrom-Json }
catch { Write-Host "Nao consegui ler o marca.json ($Marca): $($_.Exception.Message)"; exit 1 }
$empresasAtivas = @(@($cfg.empresas) | Where-Object { $_.sigla -and $_.ativa -ne $false } | ForEach-Object { "$($_.sigla)".ToUpper() })
if ($empresasAtivas.Count -eq 0) { Write-Host "marca.json: nenhuma empresa ativa."; exit 1 }
$EMPRESA_PADRAO = if (Valor $cfg "empresa_padrao") { "$($cfg.empresa_padrao)".ToUpper() } else { $empresasAtivas[0] }
$ORGANIZACAO = "$(Valor $cfg 'organizacao')"
$POLITICA = "$(Valor (Valor $cfg 'politica') 'codigo')"
if ($ORGANIZACAO -eq "[definir]") { $ORGANIZACAO = "" }
if ($POLITICA -eq "[definir]") { $POLITICA = "" }
$tiposExtra = Valor $cfg "tipos_extra"
if ($tiposExtra) {
  foreach ($p in $tiposExtra.PSObject.Properties) {
    $sig = $p.Name.ToUpper()
    if ($p.Value.variante -eq "curto") { $TIPOS_CURTOS += $sig } elseif ($p.Value.variante -eq "longo") { $TIPOS_LONGOS += $sig }
  }
}

# ---------------------------------------------------------------- varredura
$raiz = (Resolve-Path -LiteralPath $Pasta).Path
Write-Host "INVENTARIO de $raiz"

# procura um lista-mestra.json que ja exista na arvore, para nao criar uma segunda lista
$listasNaArvore = @(Get-ChildItem -LiteralPath $raiz -Recurse -File -Filter "lista-mestra.json" -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.FullName } | Sort-Object)
if ($Lista) {
  $destino = $Lista
} elseif ($listasNaArvore.Count -eq 1) {
  $destino = $listasNaArvore[0]
  Write-Host "  lista mestra encontrada em: $destino"
} elseif ($listasNaArvore.Count -gt 1) {
  Write-Host "  Ha mais de um lista-mestra.json nesta arvore:"
  foreach ($l in $listasNaArvore) { Write-Host "    $l" }
  Write-Host "  Use -Lista para dizer qual vale."
  return
} else {
  $destino = $raiz
}
if (-not $destino.ToLower().EndsWith(".json")) { $destino = Join-Path $destino "lista-mestra.json" }

# Pastas fora da varredura: Modelos guarda os arquivos de partida, que nao sao documentos. Obsoletos e lida, mas so
# decide a entrada de um codigo que nao tem mais nenhuma versao em uso. Mesmas regras do --inventariar do gerador.
$PASTAS_IGNORADAS = @("modelos")
# Documento confidencial mora numa pasta chamada Confidencial, com permissao so para quem pode ler. A pasta fica fora da
# varredura (quem nao tem acesso nao trava; quem tem nao copia nada dela para a lista), e a lista mestra, aberta a todos,
# guarda do confidencial so um esqueleto: codigo, versao, estado e data. Mesmas regras da pagina e do gerador.
$TITULO_CONFIDENCIAL = "Documento confidencial"
$FORA_DO_CONFIDENCIAL = @("responsavel", "arquivo", "projeto", "substituido_por", "plano", "encerrado_por", "rodadas_sha256")
function Eh-Confidencial($spec) { return ("$($spec['classificacao'])".Trim().ToLower() -eq "confidencial") }
$puladas = @(); $confidenciais = @()
$pular = @($Ignorar | ForEach-Object { $_.ToLower() }) + $PASTAS_IGNORADAS + @("confidencial")
# caminho relativo a raiz, sempre com "/" (como o gerador e a pagina registram a pasta)
function Relativo([string]$caminho) { return $caminho.Substring($raiz.Length).TrimStart("\", "/").Replace("\", "/") }
function Fora-Da-Varredura([string]$rel) {
  foreach ($p in @($rel -split "/" | Where-Object { $_ })) { if ($p.StartsWith(".") -or ($pular -contains $p.ToLower())) { return $true } }
  return $false
}
$visitadas = New-Object System.Collections.Generic.HashSet[string]
[void]$visitadas.Add("")
foreach ($d in Get-ChildItem -LiteralPath $raiz -Recurse -Directory -ErrorAction SilentlyContinue) {
  $r = Relativo $d.FullName
  if ($d.Name.ToLower() -eq "confidencial") { $puladas += $r }
  if (-not (Fora-Da-Varredura $r)) { [void]$visitadas.Add($r) }
}

$achados = @(); $semId = @(); $ignorados = 0; $modelos = @()
foreach ($arq in Get-ChildItem -LiteralPath $raiz -Recurse -File -ErrorAction SilentlyContinue) {
  if ($arq.Name.StartsWith("~$") -or $arq.Name.StartsWith(".")) { continue }
  $pastaRel = Relativo $arq.DirectoryName
  if (Fora-Da-Varredura $pastaRel) { continue }
  if ($arq.Extension -ne ".docx") { if ($arq.Extension -ne ".json") { $ignorados++ }; continue }
  $rel = Relativo $arq.FullName
  try { $r = Ficha-Do-Docx $arq.FullName }
  catch { $semId += ,@($rel, "nao foi possivel abrir o arquivo"); continue }
  $spec = $r.spec; $origem = $r.origem
  # arquivo de partida de modelo: nao e documento, nao entra na lista
  if ($spec.Contains("modelo") -and $spec["modelo"]) { $modelos += $rel; continue }
  $peloNome = Ficha-Do-Nome $arq.Name
  if ($origem -eq "sem_timbrado") {
    if ($null -eq $peloNome) { $semId += ,@($rel, "sem ficha, sem codigo no cabecalho e fora do padrao de nome"); continue }
    $spec = $peloNome; $origem = "nome"
  } elseif ($peloNome) {
    foreach ($k in $peloNome.Keys) { if (-not $spec.Contains($k) -or $null -eq $spec[$k] -or "$($spec[$k])" -eq "") { $spec[$k] = $peloNome[$k] } }
  }
  $sigla = Sigla-Da-Pasta $arq.Directory.Name
  if ((-not $spec.Contains("tipo") -or -not $spec["tipo"]) -and $sigla) { $spec["tipo"] = $sigla }
  if (-not $spec.Contains("tipo") -or -not $spec["tipo"]) { $semId += ,@($rel, "nao deu para saber o tipo do documento"); continue }
  if (-not $spec.Contains("empresa") -or -not $spec["empresa"]) { $spec["empresa"] = $EMPRESA_PADRAO }
  if (-not $spec.Contains("area")    -or -not $spec["area"])    { $spec["area"] = "ADM" }
  if (-not $spec.Contains("estado")  -or -not $spec["estado"])  { $spec["estado"] = "MINUTA" }
  if (-not $spec.Contains("versao")  -or -not $spec["versao"])  { $spec["versao"] = "1.0" }
  if (-not $spec.Contains("titulo")  -or -not $spec["titulo"])  { $spec["titulo"] = [IO.Path]::GetFileNameWithoutExtension($arq.Name) }
  if ((-not $spec.Contains("ano") -or -not $spec["ano"]) -and $spec.Contains("codigo") -and "$($spec['codigo'])" -match "/(\d{4})$") { $spec["ano"] = [int]$Matches[1] }
  if (-not $spec.Contains("codigo") -or -not $spec["codigo"]) {
    $num = "XXX"
    if ($spec.Contains("numero") -and $spec["numero"]) { $num = "{0:d3}" -f [int]$spec["numero"] }
    $codigo = "$($spec['empresa'])-$($spec['tipo'])-$($spec['area'])-$num"
    if (($TIPOS_CURTOS -contains $spec["tipo"]) -and $spec.Contains("ano") -and $spec["ano"]) { $codigo += "/$($spec['ano'])" }
    $spec["codigo"] = $codigo
  }
  if (-not $pastaRel) { $pastaRel = "." }
  $obsoleta = (Pasta-Obsoletos $pastaRel) -or ("$($spec['estado'])" -eq "OBSOLETO")
  $achados += ,[pscustomobject]@{ spec = $spec; origem = $origem; nome = $arq.Name; pasta = $pastaRel; obsoleta = $obsoleta }
}

# ---------------------------------------------------------------- uma entrada por codigo
$existentes = [ordered]@{}
$antiga = $null
if (Test-Path -LiteralPath $destino) {
  try {
    $antiga = (Get-Content -LiteralPath $destino -Raw -Encoding UTF8) | ConvertFrom-Json
    foreach ($d in @($antiga.documentos)) { if ($d.codigo) { $existentes[(Chave-Registro $d.codigo $d.numero $d.arquivo)] = $d } }
  } catch { Write-Host "  (lista existente ilegivel: sera refeita)" }
}

$porCodigo = [ordered]@{}
foreach ($a in $achados) {
  $num = $null; if ($a.spec.Contains("numero")) { $num = $a.spec["numero"] }
  $c = Chave-Registro $a.spec["codigo"] $num $a.nome
  if (-not $porCodigo.Contains($c)) { $porCodigo[$c] = @() }
  $porCodigo[$c] += ,$a
}

$agora = (Get-Date).ToString("dd/MM/yyyy HH:mm")
$documentos = @(); $conflitos = @(); $varias = @(); $alertas = @()
$versaoBase = { param($v) ("$v" -split "-r", 2)[0] }
foreach ($chave in @($porCodigo.Keys | Sort-Object)) {
  $itens = @($porCodigo[$chave] | Sort-Object @{ Expression = { (Chave-Versao $_.spec["versao"])[0] } },
                                              @{ Expression = { (Chave-Versao $_.spec["versao"])[1] } },
                                              @{ Expression = { (Chave-Versao $_.spec["versao"])[2] } },
                                              @{ Expression = { (Chave-Versao $_.spec["versao"])[3] } },
                                              @{ Expression = { $ORDEM_ESTADO["$($_.spec['estado'])"] } },
                                              @{ Expression = { if (Projeto-Da-Pasta $_.pasta) { 1 } else { 0 } } },
                                              @{ Expression = { $_.pasta } })
  $codigo = ($chave -split "\|", 2)[0]
  # vale a versao em uso: a oficial mais alta; sem oficial, a mais alta em uso; sem nenhuma em uso, a copia obsoleta
  $emUso = @($itens | Where-Object { -not $_.obsoleta })
  $oficiais = @($emUso | Where-Object { "$($_.spec['estado'])" -eq "OFICIAL" })
  $copias = @($itens | Where-Object { $_.obsoleta })
  if ($oficiais.Count -gt 0) { $atual = $oficiais[-1] } elseif ($emUso.Count -gt 0) { $atual = $emUso[-1] } else { $atual = $copias[-1] }
  # oficial cuja versao ja tem copia obsoleta: foi substituida e o arquivo ficou para tras na pasta
  if ("$($atual.spec['estado'])" -eq "OFICIAL") {
    $ja = $null
    foreach ($c in $copias) { if ((Compara-Versao $c.spec["versao"] $atual.spec["versao"]) -ge 0) { $ja = $c; break } }
    if ($ja) {
      $alertas += "$codigo v$($atual.spec['versao']): o arquivo oficial ainda esta em $($atual.pasta), mas essa versao ja tem copia obsoleta em $($ja.pasta). Retire o oficial da pasta."
      $atual = $ja
    }
  }
  $s = $atual.spec
  $listaOficiais = ($oficiais | ForEach-Object { "v$($_.spec['versao']) em $($_.pasta)" }) -join "; "
  if ($oficiais.Count -gt 1 -and (Tem-Rodadas $s)) {
    $alertas += "${codigo}: relatorio de conducao com mais de um arquivo em uso ($listaOficiais). A v$($oficiais[-1].spec['versao']) contem as rodadas das anteriores: retire as anteriores da pasta (relatorio de conducao nao tem copia obsoleta)."
  } elseif ($oficiais.Count -gt 1) {
    # a mesma versao em mais de uma pasta e copia do mesmo documento, nao versao anterior
    $ultV = $oficiais[-1].spec["versao"]
    $iguais = @($oficiais | Where-Object { (Compara-Versao $_.spec["versao"] $ultV) -eq 0 })
    if ($iguais.Count -lt $oficiais.Count) {
      $alertas += "${codigo}: mais de uma versao oficial em uso ($listaOficiais). A anterior precisa ser tornada obsoleta."
    }
    if ($iguais.Count -gt 1) {
      $alertas += "${codigo} v${ultV}: o mesmo documento oficial esta em mais de uma pasta ($(($iguais | ForEach-Object { $_.pasta }) -join '; ')). Deixe so uma copia: documento de projeto fica na pasta do projeto."
    }
  }
  if ($oficiais.Count -gt 0) {
    $ult = $oficiais[-1].spec["versao"]
    foreach ($m in $emUso) {
      if ("$($m.spec['estado'])" -ne "OFICIAL" -and (Compara-Versao (& $versaoBase $m.spec["versao"]) (& $versaoBase $ult)) -le 0) {
        $alertas += "${codigo}: a versao $($m.spec['versao']) ($($m.spec['estado']), em $($m.pasta)) e de uma versao ja publicada (v$ult) e pode sair da pasta."
      }
    }
  }
  foreach ($c in $copias) {
    if (-not (Pasta-Obsoletos $c.pasta)) { $alertas += "$($c.nome): copia obsoleta fora da pasta OBSOLETOS (esta em $($c.pasta))." }
  }
  foreach ($m in $emUso) {
    $pj = ""; if ($m.spec.Contains("projeto")) { $pj = "$($m.spec['projeto'])" }
    if ($pj -and (Sem-Acento (Projeto-Da-Pasta $m.pasta)) -ne (Sem-Acento $pj)) {
      $alertas += "$($m.nome): e do projeto `"$pj`", mas esta em $($m.pasta). Documento de projeto fica em Projetos/$pj."
    }
  }
  $titulos = @($itens | ForEach-Object { $_.spec["titulo"] } | Where-Object { $_ } | Select-Object -Unique)
  if (@($titulos | ForEach-Object { Sem-Acento $_ } | Select-Object -Unique).Count -gt 1) { $conflitos += ,@($codigo, $titulos) }
  if ($itens.Count -gt 1) { $varias += ,@($codigo, @($itens | ForEach-Object { "$($_.spec['versao']) $($_.spec['estado'])" })) }
  $antes = $existentes[$chave]
  $numero = $null; if ($s.Contains("numero") -and $s["numero"]) { $numero = [int]$s["numero"] }
  $ano = $null;    if ($s.Contains("ano") -and $s["ano"])       { $ano = [int]$s["ano"] }
  $data = $null
  foreach ($k in @("data_aprovacao", "data_emissao")) { if (-not $data -and $s.Contains($k) -and $s[$k]) { $data = $s[$k] } }
  $resp = ""; if ($s.Contains("elaborado_por") -and $s["elaborado_por"]) { $resp = $s["elaborado_por"] }
  $reg = [ordered]@{
    codigo = $s["codigo"]; titulo = $s["titulo"]; empresa = $s["empresa"]; tipo = $s["tipo"]; area = $s["area"]
    numero = $numero; ano = $ano; versao = $s["versao"]; estado = $s["estado"]; data = $data
    responsavel = $resp; arquivo = $atual.nome
  }
  # substituido_por so vale para a entrada obsoleta
  if ("$($s['estado'])" -eq "OBSOLETO" -and $s.Contains("substituido_por") -and $s["substituido_por"]) { $reg["substituido_por"] = $s["substituido_por"] }
  if ($s.Contains("projeto") -and $s["projeto"]) { $reg["projeto"] = $s["projeto"] }
  if (Tem-Rodadas $s) {   # relatorio de conducao: plano e impressao digital das rodadas
    $reg["plano"] = $s["plano"]
    $reg["rodadas_sha256"] = Digitais-Rodadas $s["rodadas"]
    $fim = Encerramento-Rel $s["rodadas"]
    if ($fim) { $reg["encerrado_por"] = Valor $fim "termo" }
  }
  $reg["pasta"] = $atual.pasta; $reg["origem"] = $atual.origem
  if (-not $reg.Contains("projeto") -and (Projeto-Da-Pasta $atual.pasta)) { $reg["projeto"] = Projeto-Da-Pasta $atual.pasta }   # documento antigo, sem o campo na ficha
  if ($itens.Count -gt 1) { $reg["arquivos"] = $itens.Count }
  if (Tem-Rodadas $s) {
    # as rodadas do arquivo conferem com as registradas na lista e com as das versoes anteriores ainda na pasta?
    # rodada que nao confere nao passa para a lista: ela mantem as registradas
    $hs = @($reg["rodadas_sha256"])
    $pub = @(); if ($antes -and (Valor $antes "rodadas_sha256")) { $pub = @($antes.rodadas_sha256) }
    if ($pub.Count -gt $hs.Count) {
      $alertas += "${codigo}: a lista mestra registra $($pub.Count) rodada(s), mas o arquivo mais recente na pasta (v$($s['versao'])) tem $($hs.Count). A versao vigente sumiu da pasta?"
      $reg["rodadas_sha256"] = $pub
    } elseif ($pub.Count -gt 0 -and (($hs[0..($pub.Count - 1)]) -join ",") -ne ($pub -join ",")) {
      $alertas += "$codigo v$($s['versao']): as rodadas nao conferem com as registradas na lista mestra (rodada publicada alterada). A lista mantem as rodadas registradas."
      $reg["rodadas_sha256"] = $pub
    }
    foreach ($m in $emUso) {
      if ([object]::ReferenceEquals($m.spec, $s)) { continue }
      # a funcao devolve a lista inteira como um objeto so: guardar antes de abrir em @() (senao vira lista de lista)
      $h2 = @(); if ($m.spec.Contains("rodadas")) { $t = Digitais-Rodadas $m.spec["rodadas"]; $h2 = @($t) }
      $prefixo = if ($h2.Count -gt 0) { ($hs[0..([Math]::Min($h2.Count, $hs.Count) - 1)]) -join "," } else { "" }
      if ($h2.Count -gt $hs.Count -or ($h2 -join ",") -ne $prefixo) {
        $alertas += "$($m.nome): as rodadas desta versao nao conferem com as da v$($s['versao']) (rodada publicada alterada)."
      }
    }
  }
  foreach ($i in $itens) {
    if (Eh-Confidencial $i.spec) { $alertas += ("{0}: documento confidencial fora da pasta Confidencial (esta em {1}). Quem tem acesso a pasta comum pode le-lo: mova-o para a pasta Confidencial." -f $i.nome, $i.pasta) }
  }
  $conf = Eh-Confidencial $s
  if ($conf) {
    # confidencial sem numero: nao reserva nada, e o esqueleto nao diria qual e
    if (-not $numero) { $existentes.Remove($chave) | Out-Null; continue }
    $reg = [ordered]@{
      codigo = $s["codigo"]; titulo = $TITULO_CONFIDENCIAL; empresa = $s["empresa"]; tipo = $s["tipo"]; area = $s["area"]
      numero = $numero; ano = $ano; versao = $s["versao"]; estado = $s["estado"]; data = $data; confidencial = $true
      pasta = $atual.pasta; origem = $atual.origem
    }
    if ($itens.Count -gt 1) { $reg["arquivos"] = $itens.Count }
  }
  # mesma regra do gerador: a entrada antiga e atualizada, e o que so valia para a obsoleta nao sobrevive
  $novo = [ordered]@{}
  if ($antes) { foreach ($p in $antes.PSObject.Properties) { $novo[$p.Name] = $p.Value } }
  foreach ($k in $reg.Keys) { $novo[$k] = $reg[$k] }
  foreach ($k in @("substituido_por", "motivo_obsoleto")) { if (-not $reg.Contains($k) -and $novo.Contains($k)) { $novo.Remove($k) } }
  # confidencial: o esqueleto apaga o que a entrada tinha antes; deixou de ser confidencial, a marca sai
  $tirar = if ($conf) { $FORA_DO_CONFIDENCIAL } else { @("confidencial") }
  foreach ($k in $tirar) { if (-not $reg.Contains($k) -and $novo.Contains($k)) { $novo.Remove($k) } }
  $novo["registrado_em"] = if ($antes -and $antes.registrado_em) { $antes.registrado_em } else { $agora }
  $novo["atualizado_em"] = $agora
  $documentos += ,$novo
  $existentes.Remove($chave) | Out-Null
}
# entradas da lista sem arquivo: saem quando a pasta registrada foi varrida (ou e a de Modelos); fora disso, ficam e
# sao avisadas
$removidas = @(); $naoConferidas = @()
foreach ($c in @($existentes.Keys)) {
  $d = $existentes[$c]
  $p = "$($d.pasta)".Replace("\", "/")
  if (Valor $d "confidencial") { $confidenciais += "$($d.codigo)" }   # esta na pasta Confidencial, que nao e varrida
  elseif ($null -eq $d.PSObject.Properties["pasta"]) { $naoConferidas += "$($d.codigo) (sem pasta registrada)" }
  else {
    if ($p -eq ".") { $p = "" }
    $ignorada = $false
    foreach ($x in @($p -split "/")) { if ($PASTAS_IGNORADAS -contains $x.ToLower()) { $ignorada = $true } }
    if ($ignorada -or $visitadas.Contains($p)) {
      $arqOuPasta = "$(Valor $d 'arquivo')"; if (-not $arqOuPasta) { $arqOuPasta = $p }
      $removidas += "$($d.codigo) v$($d.versao) ($arqOuPasta)"
      continue
    }
    $naoConferidas += "$($d.codigo) ($p)"
  }
  $o = [ordered]@{}
  foreach ($prop in $d.PSObject.Properties) { $o[$prop.Name] = $prop.Value }
  $documentos += ,$o
}
$documentos = @($documentos | Sort-Object @{ Expression = { "$($_.empresa)" } }, @{ Expression = { "$($_.tipo)" } },
                                          @{ Expression = { "$($_.area)" } }, @{ Expression = { "$($_.ano)" } },
                                          @{ Expression = { [int]("0" + "$($_.numero)") } }, @{ Expression = { "$($_.arquivo)" } })
# relatorio de conducao: o plano dele tem de estar oficial
foreach ($d in $documentos) {
  if ("$($d.tipo)" -eq "REL" -and $d.plano) {
    $pl = @($documentos | Where-Object { "$($_.codigo)" -eq "$($d.plano)" })
    if ($pl.Count -eq 0 -or "$($pl[0].estado)" -ne "OFICIAL") { $alertas += "$($d.codigo): o plano $($d.plano) nao esta oficial na lista mestra." }
  }
}

# ---------------------------------------------------------------- relatorio
$conta = { param($valores)
  $t = ($valores | Group-Object | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) $($_.Count)" }) -join " - "
  if ($t) { $t } else { "nenhum" } }
Write-Host ("  {0} arquivo(s) .docx  -  {1} documento(s) distintos" -f $achados.Count, $porCodigo.Count)
Write-Host ("  origem dos dados: " + (& $conta @($achados | ForEach-Object { $_.origem })))
Write-Host ("  por tipo: "         + (& $conta @($achados | ForEach-Object { $_.spec["tipo"] })))
Write-Host ("  por estado: "       + (& $conta @($achados | ForEach-Object { $_.spec["estado"] })))
Write-Host ""
foreach ($d in $documentos) {
  $marca = " "; if ($d.origem -eq "nome" -or $d.origem -eq "cabecalho") { $marca = "*" }
  $resp = "$($d.responsavel)"; if (-not $resp) { $resp = "-" }
  $titulo = "$($d.titulo)"; if ($titulo.Length -gt 44) { $titulo = $titulo.Substring(0, 44) }
  Write-Host ("{0} {1,-22} v{2,-8} {3,-10} {4,-44} {5,-10} {6}" -f $marca, $d.codigo, $d.versao, $d.estado, $titulo, $d.pasta, $resp)
}
Write-Host ""
if (@($documentos | Where-Object { $_.origem -eq "nome" -or $_.origem -eq "cabecalho" }).Count -gt 0) {
  Write-Host "  (* dados lidos do cabecalho ou do nome do arquivo: confira titulo e responsavel)"
}
foreach ($al in $alertas) { Write-Host "  ATENCAO - $al" }
if ($removidas.Count -gt 0) {
  Write-Host ("  Sairam da lista {0} entrada(s) cujo arquivo nao esta mais na pasta:" -f $removidas.Count)
  foreach ($x in $removidas) { Write-Host "    $x" }
}
if ($puladas.Count -gt 0) { Write-Host ("  Pasta Confidencial fora da varredura ({0}): o que esta nela nao entra na lista, nem o titulo." -f ($puladas -join "; ")) }
if ($confidenciais.Count -gt 0) { Write-Host ("  {0} documento(s) confidencial(is) na lista ficaram sem conferir (estao na pasta Confidencial): {1}" -f $confidenciais.Count, ($confidenciais -join "; ")) }
if ($naoConferidas.Count -gt 0) {
  Write-Host ("  {0} entrada(s) ficaram sem conferir (pasta nao varrida): {1}" -f $naoConferidas.Count, ($naoConferidas -join "; "))
}
if ($varias.Count -gt 0) {
  Write-Host ("  {0} documento(s) com mais de um arquivo (na lista ficou a versao em uso: a oficial mais alta, quando ha):" -f $varias.Count)
  foreach ($v in $varias) { Write-Host ("    {0}: {1}" -f $v[0], ($v[1] -join ", ")) }
}
if ($conflitos.Count -gt 0) {
  Write-Host "  ATENCAO - mesmo codigo com titulos diferentes (numeracao repetida):"
  foreach ($c in $conflitos) { Write-Host ("    {0}: {1}" -f $c[0], ($c[1] -join " | ")) }
}
$versoesRuins = @($achados | Where-Object { "$($_.spec['versao'])" -notmatch "^\d+\.\d+(-r\d+)?$" } |
                   ForEach-Object { "$($_.spec['codigo']): v$($_.spec['versao'])" } | Select-Object -Unique)
if ($versoesRuins.Count -gt 0) {
  Write-Host "  Versoes fora do padrao X.Y ou X.Y-rN (acerte ao gerar a proxima versao):"
  foreach ($v in $versoesRuins) { Write-Host "    $v" }
}
$semNumero = @($achados | Where-Object { -not $_.spec.Contains("numero") -or -not $_.spec["numero"] })
if ($semNumero.Count -gt 0) {
  Write-Host ("  {0} arquivo(s) sem numero (codigo com XXX): reserve o numero:" -f $semNumero.Count)
  foreach ($x in $semNumero) { Write-Host ("    {0} - {1}" -f $x.spec["codigo"], $x.nome) }
}
if ($semId.Count -gt 0) {
  Write-Host ("  {0} arquivo(s) .docx nao identificados (ficaram de fora):" -f $semId.Count)
  foreach ($s in $semId) { Write-Host ("    {0} - {1}" -f $s[0], $s[1]) }
}
if ($ignorados -gt 0) { Write-Host ("  {0} arquivo(s) que nao sao .docx foram ignorados." -f $ignorados) }
if ($modelos.Count -gt 0) { Write-Host ("  {0} arquivo(s) de partida de modelo foram pulados (nao sao documentos)." -f $modelos.Count) }
Write-Host ""

if ($Conferir) { Write-Host "Nada foi gravado (-Conferir)."; return }

# nome diferente de $Lista (parametro): variavel de mesmo nome herdaria o tipo [string] do parametro
# lista que ja existia mantem o nome e o padrao dela; lista nova leva os do marca.json
$nomeLista = $ORGANIZACAO; if (Valor $antiga "lista_mestra") { $nomeLista = $antiga.lista_mestra }
$padraoLista = $POLITICA;  if (Valor $antiga "padrao") { $padraoLista = $antiga.padrao }
$controle = [ordered]@{
  lista_mestra = $nomeLista; padrao = $padraoLista
  atualizado_em = $agora; documentos = $documentos
}
$json = ConvertTo-Json -InputObject $controle -Depth 6   # -InputObject (e nao pipe): pipe converte o dicionario em texto
$pastaDestino = [IO.Path]::GetDirectoryName($destino)
if ($pastaDestino -and -not (Test-Path -LiteralPath $pastaDestino)) { New-Item -ItemType Directory -Path $pastaDestino | Out-Null }
[IO.File]::WriteAllText($destino, $json + "`n", (New-Object System.Text.UTF8Encoding $false))
Write-Host ("LISTA MESTRA: {0} ({1} documentos)" -f $destino, $documentos.Count)
