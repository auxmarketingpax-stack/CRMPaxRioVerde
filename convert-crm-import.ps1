param(
  [string]$CrmWorkbookPath = "\\prv-cml01\pessoal\wendller.ferreira\Downloads\Controle de CRM 2026 (1).xlsx",
  [string]$MarketingWorkbookPath = "\\prv-cml01\pessoal\wendller.ferreira\Downloads\Controle de Contratos Gerados pelo Marketing 2026 (1).xlsx",
  [string]$ReferralWorkbookPath = "\\prv-cml01\pessoal\wendller.ferreira\Downloads\PROGRAMA DE INDICAÇÕES COLABORADOR.xlsx",
  [string]$OutputDir = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $OutputDir) {
  $OutputDir = Join-Path $PSScriptRoot "imports_crm"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$LabelIndicacao = "Indica$([char]0x00E7)$([char]0x00E3)o"
$LabelIndicacaoColaborador = "$LabelIndicacao Colaborador"
$LabelApresentacaoNegociacao = "Apresenta$([char]0x00E7)$([char]0x00E3)o/Negocia$([char]0x00E7)$([char]0x00E3)o"
$LabelNaoResponde = "N$([char]0x00E3)o responde"
$LabelNaoTemInteresse = "N$([char]0x00E3)o tem interesse"
$LabelJaEAssociado = "J$([char]0x00E1) $([char]0x00E9) associado"
$LabelClienteNaoPodeReativar = "Cliente n$([char]0x00E3)o pode reativar o plano"

function Normalize-Whitespace {
  param([AllowNull()][string]$Value)

  if ($null -eq $Value) { return "" }
  $text = [string]$Value
  $text = $text.Replace([char]160, " ")
  $text = [regex]::Replace($text, "\s+", " ")
  return $text.Trim()
}

function Remove-Diacritics {
  param([AllowNull()][string]$Value)

  $text = Normalize-Whitespace $Value
  if (-not $text) { return "" }

  $normalized = $text.Normalize([Text.NormalizationForm]::FormD)
  $builder = New-Object System.Text.StringBuilder
  foreach ($char in $normalized.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$builder.Append($char)
    }
  }
  return $builder.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Normalize-LookupText {
  param([AllowNull()][string]$Value)

  return (Remove-Diacritics $Value).ToLowerInvariant()
}

function Clean-Phone {
  param([AllowNull()][string]$Value)

  $digits = [regex]::Replace((Normalize-Whitespace $Value), "\D", "")
  if (-not $digits) { return "" }

  if ($digits.Length -ge 12 -and $digits.StartsWith("55")) {
    $digits = $digits.Substring(2)
  }

  if ($digits.Length -eq 8 -or $digits.Length -eq 9) {
    $digits = "64$digits"
  }

  if ($digits.Length -lt 10) {
    return ""
  }

  if ($digits.Length -gt 11) {
    $digits = $digits.Substring($digits.Length - 11)
  }

  return $digits
}

function Convert-DateToIso {
  param([AllowNull()][string]$Value)

  $text = Normalize-Whitespace $Value
  if (-not $text) { return "" }

  $formats = @(
    "dd/MM/yyyy",
    "d/M/yyyy",
    "dd/MM/yy",
    "d/M/yy",
    "yyyy-MM-dd",
    "yyyy/MM/dd",
    "dd-MM-yyyy",
    "d-M-yyyy"
  )

  foreach ($format in $formats) {
    $date = [datetime]::MinValue
    if ([datetime]::TryParseExact($text, $format, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
      return $date.ToString("yyyy-MM-dd")
    }
  }

  $fallback = [datetime]::MinValue
  if ([datetime]::TryParse($text, [ref]$fallback)) {
    return $fallback.ToString("yyyy-MM-dd")
  }

  return ""
}

function Find-DateInText {
  param([AllowNull()][string]$Value, [int]$DefaultYear = 0)

  $text = Normalize-Whitespace $Value
  if (-not $text) { return "" }

  $matches = [regex]::Matches($text, '\b(\d{1,2})[\/\.-](\d{1,2})(?:[\/\.-](\d{2,4}))?\b')
  foreach ($match in $matches) {
    $day = [int]$match.Groups[1].Value
    $month = [int]$match.Groups[2].Value
    $yearText = $match.Groups[3].Value
    $year = 0

    if ($yearText) {
      $year = [int]$yearText
      if ($year -lt 100) {
        $year += 2000
      }
    } elseif ($DefaultYear -gt 0) {
      $year = $DefaultYear
    } else {
      continue
    }

    try {
      return (Get-Date -Year $year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0).ToString("yyyy-MM-dd")
    }
    catch {
      continue
    }
  }

  return ""
}

function Parse-MonthReference {
  param([AllowNull()][string]$Value)

  $text = Normalize-LookupText $Value
  if (-not $text) {
    return [pscustomobject]@{ Month = 0; Year = 0 }
  }

  $monthMap = @{
    "jan" = 1; "janeiro" = 1
    "fev" = 2; "fevereiro" = 2
    "mar" = 3; "marco" = 3
    "abr" = 4; "abril" = 4
    "mai" = 5; "maio" = 5
    "jun" = 6; "junho" = 6
    "jul" = 7; "julho" = 7
    "ago" = 8; "agosto" = 8
    "set" = 9; "setembro" = 9
    "out" = 10; "outubro" = 10
    "nov" = 11; "novembro" = 11
    "dez" = 12; "dezembro" = 12
  }

  $month = 0
  foreach ($key in $monthMap.Keys) {
    if ($text -match ("(^|\W)" + [regex]::Escape($key) + "(\W|$)")) {
      $month = [int]$monthMap[$key]
      break
    }
  }

  if (-not $month) {
    $numericMonth = [regex]::Match($text, '\b(0?[1-9]|1[0-2])\/(20\d{2})\b')
    if ($numericMonth.Success) {
      return [pscustomobject]@{
        Month = [int]$numericMonth.Groups[1].Value
        Year = [int]$numericMonth.Groups[2].Value
      }
    }
  }

  $yearMatch = [regex]::Match($text, '\b(20\d{2})\b')
  $year = if ($yearMatch.Success) { [int]$yearMatch.Groups[1].Value } else { 0 }

  return [pscustomobject]@{
    Month = $month
    Year = $year
  }
}

function Get-FallbackStartDate {
  param($Row)

  if (Normalize-Whitespace $Row.data_inicio) {
    return $Row.data_inicio
  }

  $sourceKey = Normalize-LookupText "$($Row.source_file) $($Row.source_sheet) $($Row.source_group)"
  $noteDate = Find-DateInText $Row.observacoes 0
  if ($noteDate) {
    return $noteDate
  }

  if ($sourceKey -like "*crm (empresas)*") {
    return "2024-01-01"
  }

  if ($sourceKey -like "*indicacoes colaborador*") {
    return "2025-02-01"
  }

  return "2026-01-01"
}

function Resolve-WorkbookPath {
  param([string]$Path)

  if (Test-Path -LiteralPath $Path) {
    return (Get-Item -LiteralPath $Path).FullName
  }

  $parent = Split-Path -Path $Path -Parent
  $leaf = Split-Path -Path $Path -Leaf
  if (-not (Test-Path -LiteralPath $parent)) {
    throw "Diretorio nao encontrado: $parent"
  }

  $targetKey = Normalize-LookupText $leaf
  $officeFiles = Get-ChildItem -LiteralPath $parent -File |
    Where-Object { @(".xlsx", ".xls", ".csv").Contains($_.Extension.ToLowerInvariant()) }

  $candidate = $officeFiles |
    Where-Object { Normalize-LookupText $_.Name -eq $targetKey } |
    Select-Object -First 1

  if (-not $candidate) {
    $needle = ($targetKey -replace "\.(xlsx|xls|csv)$", "")
    $tokens = @($needle -split "\s+" | Where-Object { $_ -and $_.Length -ge 4 })
    $candidate = $officeFiles |
      Where-Object {
        $nameKey = Normalize-LookupText $_.Name
        $nameKey -like ("*" + $needle + "*") -or
        (($tokens | Where-Object { $nameKey -like ("*" + $_ + "*") }).Count -ge [Math]::Min(2, [Math]::Max(1, $tokens.Count)))
      } |
      Select-Object -First 1
  }

  if (-not $candidate) {
    throw "Arquivo nao encontrado: $Path"
  }

  return $candidate.FullName
}

function Get-PipelineFromText {
  param([AllowNull()][string[]]$Texts)

  $joined = (($Texts | ForEach-Object { Normalize-Whitespace $_ }) -join " ").Trim()
  $lookup = Normalize-LookupText $joined

  if (-not $lookup) { return "Follow-Up" }

  if (
    $lookup -match "\bfechado\b" -or
    $lookup -match "\bfechou\b" -or
    $lookup -match "contratos empresariais" -or
    $lookup -match "\bfeito\b" -or
    $lookup -match "fez o contrato" -or
    $lookup -match "cliente diz que fez" -or
    $lookup -match "fez pela empresa" -or
    $lookup -match "realizou o plano" -or
    $lookup -match "ja fez o plano"
  ) {
    return "Fechado"
  }

  if (
    $lookup -match "nao pode reativa" -or
    $lookup -match "nao pode reativar" -or
    $lookup -match "nao consegue reativar" -or
    $lookup -match "nao seria possivel reativar"
  ) {
    return $LabelClienteNaoPodeReativar
  }

  if (
    $lookup -match "ja e cliente" -or
    $lookup -match "ja e associado" -or
    $lookup -match "cliente da pax social" -or
    $lookup -match "ja esta associado" -or
    $lookup -match "ja tem plano" -or
    $lookup -match "ja possui plano"
  ) {
    return $LabelJaEAssociado
  }

  if (
    $lookup -match "nao tem interesse" -or
    $lookup -match "nao tem interresse" -or
    $lookup -match "sem interesse" -or
    $lookup -match "sem interresse" -or
    $lookup -match "cliente sem interesse" -or
    $lookup -match "nao quer" -or
    $lookup -match "nao quer saber do plano" -or
    $lookup -match "perdido" -or
    $lookup -match "encerrado" -or
    $lookup -match "ja foi cliente" -or
    $lookup -match "mudou para outra cidade" -or
    $lookup -match "mudou para santa helena" -or
    $lookup -match "nao mora rv" -or
    $lookup -match "nao mora rio verde" -or
    $lookup -match "diretoria barrou" -or
    $lookup -match "barrou a proposta" -or
    $lookup -match "nao sendo possivel fazer o plano" -or
    $lookup -match "nao e possivel fazer o plano" -or
    $lookup -match "nao foi possivel fazer o plano"
  ) {
    return $LabelNaoTemInteresse
  }

  if (
    $lookup -match "nao responde" -or
    $lookup -match "nao reponde" -or
    $lookup -match "nao atende" -or
    $lookup -match "sem resposta" -or
    $lookup -match "nao visualiza" -or
    $lookup -match "nao responde as mensagens" -or
    $lookup -match "numero errado" -or
    $lookup -match "n errado" -or
    $lookup -match "telefone inexistente" -or
    $lookup -match "numero de telefone ixesistente" -or
    $lookup -match "numero nao chama" -or
    $lookup -match "nao possui whatsapp" -or
    $lookup -match "nao tem whatsapp" -or
    $lookup -match "nao tem zap" -or
    $lookup -match "nao possui zap" -or
    $lookup -match "ultimo contato" -or
    $lookup -match "ultima tentativa" -or
    $lookup -match "feito varias tentativas" -or
    $lookup -match "varias tentativas" -or
    $lookup -match "caixa postal" -or
    $lookup -match "mensagens temporarias"
  ) {
    return $LabelNaoResponde
  }

  if (
    $lookup -match "proposta" -or
    $lookup -match "apresentad" -or
    $lookup -match "apresentar" -or
    $lookup -match "reuniao" -or
    $lookup -match "marcar um dia" -or
    $lookup -match "marcar um horario" -or
    $lookup -match "marcar horario" -or
    $lookup -match "agend" -or
    $lookup -match "vir ao escritorio" -or
    $lookup -match "quer fazer" -or
    $lookup -match "quer fechar" -or
    $lookup -match "vai fechar" -or
    $lookup -match "vai analisar" -or
    $lookup -match "indeciso" -or
    $lookup -match "interessado" -or
    $lookup -match "interesse" -or
    $lookup -match "interresse" -or
    $lookup -match "ficou de acertar" -or
    $lookup -match "ficou de ver" -or
    $lookup -match "ficou de falar" -or
    $lookup -match "avaliando a possibilidade" -or
    $lookup -match "priorizar outros assuntos"
  ) {
    return $LabelApresentacaoNegociacao
  }

  if (
    $lookup -match "em andamento" -or
    $lookup -match "cliente esta com o vendedor" -or
    $lookup -match "cliente esta com vendedor" -or
    $lookup -match "liguei" -or
    $lookup -match "\bligar\b" -or
    $lookup -match "\bmsg\b" -or
    $lookup -match "mensagem" -or
    $lookup -match "retorno" -or
    $lookup -match "retornar" -or
    $lookup -match "aguardando resposta" -or
    $lookup -match "aguardando retorno" -or
    $lookup -match "aguardando" -or
    $lookup -match "vai falar com" -or
    $lookup -match "enviar e-mail" -or
    $lookup -match "enviado e-mail" -or
    $lookup -match "material explicativo" -or
    $lookup -match "vamos entrar em contato" -or
    $lookup -match "vamos ligar" -or
    $lookup -match "ligar na semana" -or
    $lookup -match "ligar segunda" -or
    $lookup -match "passou o contato" -or
    $lookup -match "entrar em contato"
  ) {
    return "Follow-Up"
  }

  if ($lookup -match "caixa postal") {
    return $LabelNaoResponde
  }

  if (
    $lookup -match "caixa postal"
  ) {
    return "Não responde"
  }

  return "Follow-Up"
}

function Get-LeadSourceFromText {
  param(
    [string]$DefaultSource,
    [AllowNull()][string[]]$Texts
  )

  $joined = (($Texts | ForEach-Object { Normalize-Whitespace $_ }) -join " ").Trim()
  $lookup = Normalize-LookupText $joined

  if ($lookup -match "indicac" -or $lookup -match "quem indicou") {
    return $LabelIndicacao
  }

  if (
    $lookup -match "indicac" -or
    $lookup -match "quem indicou"
  ) {
    return "Indicação"
  }

  if (
    $lookup -match "evento externo" -or
    $lookup -match "acao externa" -or
    $lookup -match "dia das maes" -or
    $lookup -match "troca de figurinha" -or
    $lookup -match "figurinhas" -or
    $lookup -match "otovive"
  ) {
    return "Evento Externo"
  }

  if (
    $lookup -match "trafego pago" -or
    $lookup -match "\bpago\b"
  ) {
    return "Pago"
  }

  return $DefaultSource
}

function Get-ChannelFromText {
  param(
    [string]$DefaultChannel,
    [AllowNull()][string[]]$Texts
  )

  $joined = (($Texts | ForEach-Object { Normalize-Whitespace $_ }) -join " ").Trim()
  $lookup = Normalize-LookupText $joined

  if (
    $lookup -match "\bcdl\b"
  ) {
    return "CDL"
  }

  if (
    $lookup -match "sudoexpo"
  ) {
    return "SUDOEXPO"
  }

  if (
    $lookup -match "prospecc" -or
    $lookup -match "prospec" -or
    $lookup -match "visita" -or
    $lookup -match "porta a porta" -or
    $lookup -match "cliente na rua" -or
    $lookup -match "field sales"
  ) {
    return "Field Sales"
  }

  return $DefaultChannel
}

function Quote-CsvValue {
  param([AllowNull()]$Value)

  $text = [string]$Value
  if ($text.IndexOfAny([char[]]@('"', ';', "`r", "`n")) -ge 0) {
    return '"' + $text.Replace('"', '""') + '"'
  }
  return $text
}

function Write-SemicolonCsv {
  param(
    [string]$Path,
    [object[]]$Rows,
    [string[]]$Columns
  )

  $lines = New-Object System.Collections.Generic.List[string]
  $lines.Add(($Columns | ForEach-Object { Quote-CsvValue $_ }) -join ";")

  foreach ($row in $Rows) {
    $line = foreach ($column in $Columns) {
      Quote-CsvValue ($row.$column)
    }
    $lines.Add(($line -join ";"))
  }

  $content = [string]::Join("`r`n", $lines)
  $encoding = New-Object System.Text.UTF8Encoding($true)
  [IO.File]::WriteAllText($Path, $content, $encoding)
}

function Build-Observation {
  param([string[]]$Texts)

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($text in $Texts) {
    $value = Normalize-Whitespace $text
    if (-not $value) { continue }
    if (-not $parts.Contains($value)) {
      $parts.Add($value)
    }
  }
  return ($parts -join " | ")
}

function Get-UniqueTextList {
  param([string[]]$Texts)

  $result = New-Object System.Collections.Generic.List[string]
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'

  foreach ($text in $Texts) {
    $value = Normalize-Whitespace $text
    if (-not $value) { continue }
    $key = Normalize-LookupText $value
    if ($seen.Add($key)) {
      $result.Add($value) | Out-Null
    }
  }

  return @($result)
}

function Get-ObservationFragments {
  param([string]$Text)

  $value = Normalize-Whitespace $Text
  if (-not $value) { return @() }

  $patterns = @(
    'Detalhes:\s*([^|]+)',
    'Status 1:\s*([^|]+)',
    'Status original:\s*([^|]+)',
    'Status do atendimento:\s*([^|]+)',
    'Status final:\s*([^|]+)',
    'Motivo nao fechamento:\s*([^|]+)'
  )

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($pattern in $patterns) {
    foreach ($match in [regex]::Matches($value, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
      $parts.Add((Normalize-Whitespace $match.Groups[1].Value)) | Out-Null
    }
  }

  if ($parts.Count) {
    return @(Get-UniqueTextList $parts.ToArray())
  }

  $segments = $value -split '\s*\|\s*'
  $kept = New-Object System.Collections.Generic.List[string]
  foreach ($segment in $segments) {
    $piece = Normalize-Whitespace $segment
    if (-not $piece) { continue }
    $key = Normalize-LookupText $piece
    if (
      $key.StartsWith("origem:") -or
      $key.StartsWith("canal original:") -or
      $key.StartsWith("cidade:") -or
      $key.StartsWith("contato pessoa:") -or
      $key.StartsWith("mes referencia:") -or
      $key.StartsWith("dia indicacao:") -or
      $key.StartsWith("lider:") -or
      $key.StartsWith("quantidade original:") -or
      $key.StartsWith("origem detalhada:") -or
      $key.StartsWith("data fechamento:") -or
      $key.StartsWith("ja e cliente:")
    ) {
      continue
    }
    $kept.Add($piece) | Out-Null
  }

  return @(Get-UniqueTextList $kept.ToArray())
}

function Clean-ObservationText {
  param([string]$Text)

  $groups = [regex]::Split((Normalize-Whitespace $Text), '\s*\|\|\s*')
  $fragments = New-Object System.Collections.Generic.List[string]
  foreach ($group in $groups) {
    foreach ($item in (Get-ObservationFragments $group)) {
      $fragments.Add($item) | Out-Null
    }
  }

  return ((Get-UniqueTextList $fragments.ToArray()) -join " | ")
}

function New-ImportRow {
  param(
    [string]$Nome,
    [string]$Contato,
    [string]$Responsavel,
    [string]$Valor,
    [string]$DataInicio,
    [string]$RedeSocial,
    [string]$Origem,
    [string]$IndicadoPor = "",
    [string]$SetorIndicado = "",
    [string]$Plano,
    [string]$NumeroContrato = "",
    [string]$Pipeline,
    [string]$Observacoes,
    [string]$SourceFile,
    [string]$SourceSheet,
    [string]$SourceSection,
    [int]$SourceRow
  )

  [pscustomobject]@{
    nome         = Normalize-Whitespace $Nome
    contato      = Clean-Phone $Contato
    responsavel  = Normalize-Whitespace $Responsavel
    valor        = Normalize-Whitespace $Valor
    data_inicio  = Convert-DateToIso $DataInicio
    rede_social  = Normalize-Whitespace $RedeSocial
    origem       = Normalize-Whitespace $Origem
    indicado_por = Normalize-Whitespace $IndicadoPor
    setor_indicado = Normalize-Whitespace $SetorIndicado
    plano        = Normalize-Whitespace $Plano
    numero_contrato = Normalize-Whitespace $NumeroContrato
    pipeline     = Normalize-Whitespace $Pipeline
    observacoes  = Normalize-Whitespace $Observacoes
    source_file  = $SourceFile
    source_sheet = $SourceSheet
    source_group = $SourceSection
    source_row   = $SourceRow
  }
}

function Add-SkippedRow {
  param(
    [System.Collections.Generic.List[object]]$Target,
    [string]$Reason,
    [string]$SourceFile,
    [string]$SourceSheet,
    [string]$SourceSection,
    [int]$SourceRow,
    [string]$RawName,
    [string]$RawContact,
    [string]$RawStatus
  )

  $Target.Add([pscustomobject]@{
    motivo       = $Reason
    source_file  = $SourceFile
    source_sheet = $SourceSheet
    source_group = $SourceSection
    source_row   = $SourceRow
    nome         = Normalize-Whitespace $RawName
    contato      = Normalize-Whitespace $RawContact
    status       = Normalize-Whitespace $RawStatus
  }) | Out-Null
}

function Has-MeaningfulText {
  param([string[]]$Values)

  foreach ($value in $Values) {
    if (Normalize-Whitespace $value) {
      return $true
    }
  }
  return $false
}

function Get-CellText {
  param($Worksheet, [int]$Row, [int]$Column)

  return [string]$Worksheet.Cells.Item($Row, $Column).Text
}

function Get-PipelineRank {
  param([string]$Pipeline)

  switch ($Pipeline) {
    "Fechado" { return 5 }
    { $_ -eq $LabelApresentacaoNegociacao } { return 4 }
    "Follow-Up" { return 3 }
    { $_ -eq $LabelNaoResponde } { return 2 }
    { $_ -eq $LabelNaoTemInteresse } { return 1 }
    { $_ -eq $LabelJaEAssociado } { return 1 }
    { $_ -eq $LabelClienteNaoPodeReativar } { return 1 }
    default { return 0 }
  }
}

function Finalize-ImportRow {
  param($Row)

  $cleanObservation = Clean-ObservationText $Row.observacoes

  if ((Normalize-LookupText $Row.origem) -eq "indicacao") {
    $Row.origem = $LabelIndicacao
  }

  if ((Normalize-LookupText $Row.rede_social) -eq "indicacao colaborador") {
    $Row.rede_social = $LabelIndicacaoColaborador
  }

  $channelKey = Normalize-LookupText $Row.rede_social
  if ($channelKey -in @("field sales", "sudoexpo", "cdl")) {
    $Row.origem = "Evento Externo"
  }

  if (-not (Normalize-Whitespace $Row.responsavel)) {
    $Row.responsavel = "LEIDIMAR"
  }

  $Row.observacoes = $cleanObservation
  $Row.pipeline = Get-PipelineFromText @($Row.pipeline, $cleanObservation)
  return $Row
}

$CrmWorkbookPath = Resolve-WorkbookPath $CrmWorkbookPath
$MarketingWorkbookPath = Resolve-WorkbookPath $MarketingWorkbookPath
$referralParent = Split-Path -Path $ReferralWorkbookPath -Parent
$ReferralWorkbookPath = (
  Get-ChildItem -LiteralPath $referralParent -File -Filter "*.xlsx" |
    Where-Object {
      $nameKey = Normalize-LookupText $_.Name
      $nameKey -like "*programa*" -and
      $nameKey -like "*indic*" -and
      $nameKey -like "*colaborador*"
    } |
    Select-Object -First 1 -ExpandProperty FullName
)

if (-not $ReferralWorkbookPath) {
  throw "Arquivo da planilha de indicacoes nao encontrado em $referralParent"
}

$allRows = New-Object System.Collections.Generic.List[object]
$skippedRows = New-Object System.Collections.Generic.List[object]
$duplicateRows = New-Object System.Collections.Generic.List[object]

$excel = $null
$crmWorkbook = $null
$marketingWorkbook = $null
$referralWorkbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $crmWorkbook = $excel.Workbooks.Open($CrmWorkbookPath, 0, $true)

  $crmSheet = $crmWorkbook.Worksheets.Item("CRM (Empresas)")
  $crmUsedRows = $crmSheet.UsedRange.Rows.Count
  for ($row = 2; $row -le $crmUsedRows; $row++) {
    $empresa = Get-CellText $crmSheet $row 2
    $contato = Get-CellText $crmSheet $row 3
    $nomeContato = Get-CellText $crmSheet $row 4
    $formaProspeccao = Get-CellText $crmSheet $row 5
    $cidade = Get-CellText $crmSheet $row 6
    $dataContato = Get-CellText $crmSheet $row 7
    $status1 = Get-CellText $crmSheet $row 8
    $status = Get-CellText $crmSheet $row 9
    $temInteresse = Get-CellText $crmSheet $row 10

    if (-not (Has-MeaningfulText @($empresa, $contato, $nomeContato, $formaProspeccao, $status1, $status, $temInteresse))) {
      continue
    }

    $cleanContact = Clean-Phone $contato
    if (-not (Normalize-Whitespace $empresa)) {
      Add-SkippedRow $skippedRows "Linha sem nome da empresa." (Split-Path $CrmWorkbookPath -Leaf) "CRM (Empresas)" "Empresas" $row $empresa $contato "$status / $temInteresse"
      continue
    }
    if (-not $cleanContact) {
      Add-SkippedRow $skippedRows "Contato vazio ou invalido." (Split-Path $CrmWorkbookPath -Leaf) "CRM (Empresas)" "Empresas" $row $empresa $contato "$status / $temInteresse"
      continue
    }

    $obs = Build-Observation @($temInteresse, $status1, $status)

    $allRows.Add((New-ImportRow `
      -Nome $empresa `
      -Contato $contato `
      -Responsavel "" `
      -Valor "" `
      -DataInicio $dataContato `
      -RedeSocial (Get-ChannelFromText $formaProspeccao @($formaProspeccao, $status1, $status, $temInteresse)) `
      -Origem (Get-LeadSourceFromText "Organico" @($formaProspeccao, $status1, $status, $temInteresse)) `
      -Plano "Plano Empresarial" `
      -Pipeline (Get-PipelineFromText @($status, $status1, $temInteresse)) `
      -Observacoes $obs `
      -SourceFile (Split-Path $CrmWorkbookPath -Leaf) `
      -SourceSheet "CRM (Empresas)" `
      -SourceSection "Empresas" `
      -SourceRow $row)) | Out-Null
  }

  $leads2025Sheet = $crmWorkbook.Worksheets.Item("Leads 2025")
  $leads2025Rows = $leads2025Sheet.UsedRange.Rows.Count
  for ($row = 2; $row -le $leads2025Rows; $row++) {
    $nome = Get-CellText $leads2025Sheet $row 1
    $contato = Get-CellText $leads2025Sheet $row 2
    $jaCliente = Get-CellText $leads2025Sheet $row 3
    $tipoPlano = Get-CellText $leads2025Sheet $row 4
    $statusAtendimento = Get-CellText $leads2025Sheet $row 5
    $tipoContato = Get-CellText $leads2025Sheet $row 6
    $data = Get-CellText $leads2025Sheet $row 7
    $status = Get-CellText $leads2025Sheet $row 8

    if (-not (Has-MeaningfulText @($nome, $contato, $jaCliente, $tipoPlano, $statusAtendimento, $tipoContato, $status))) {
      continue
    }

    $cleanContact = Clean-Phone $contato
    if (-not (Normalize-Whitespace $nome)) {
      Add-SkippedRow $skippedRows "Linha sem nome do lead." (Split-Path $CrmWorkbookPath -Leaf) "Leads 2025" "Leads 2025" $row $nome $contato $status
      continue
    }
    if (-not $cleanContact) {
      Add-SkippedRow $skippedRows "Contato vazio ou invalido." (Split-Path $CrmWorkbookPath -Leaf) "Leads 2025" "Leads 2025" $row $nome $contato $status
      continue
    }

    $obs = Build-Observation @($statusAtendimento, $status, $jaCliente)

    $allRows.Add((New-ImportRow `
      -Nome $nome `
      -Contato $contato `
          -Responsavel "" `
          -Valor "" `
          -DataInicio $data `
      -RedeSocial $tipoContato `
          -Origem (Get-LeadSourceFromText "Organico" @($tipoContato, $statusAtendimento, $status)) `
      -Plano $tipoPlano `
      -Pipeline (Get-PipelineFromText @($status, $statusAtendimento, $jaCliente)) `
      -Observacoes $obs `
      -SourceFile (Split-Path $CrmWorkbookPath -Leaf) `
      -SourceSheet "Leads 2025" `
      -SourceSection "Leads 2025" `
      -SourceRow $row)) | Out-Null
  }

  $crmWorkbook.Close($false)
  $crmWorkbook = $null

  $marketingWorkbook = $excel.Workbooks.Open($MarketingWorkbookPath, 0, $true)
  $vendasSheet = $marketingWorkbook.Worksheets.Item("Vendas")
  $vendasRows = $vendasSheet.UsedRange.Rows.Count

  for ($row = 5; $row -le $vendasRows; $row++) {
    $orgData = Get-CellText $vendasSheet $row 1
    $orgQuantidade = Get-CellText $vendasSheet $row 2
    $orgNome = Get-CellText $vendasSheet $row 3
    $orgContato = Get-CellText $vendasSheet $row 4
    $orgRede = Get-CellText $vendasSheet $row 5
    $orgVendedor = Get-CellText $vendasSheet $row 6
    $orgStatus = Get-CellText $vendasSheet $row 7

    if (Has-MeaningfulText @($orgData, $orgQuantidade, $orgNome, $orgContato, $orgRede, $orgVendedor, $orgStatus)) {
      $cleanContact = Clean-Phone $orgContato
      if ((Normalize-Whitespace $orgNome) -and $cleanContact) {
        $obs = Build-Observation @($orgStatus)

        $allRows.Add((New-ImportRow `
          -Nome $orgNome `
          -Contato $orgContato `
          -Responsavel $orgVendedor `
          -Valor "" `
          -DataInicio $orgData `
          -RedeSocial $orgRede `
          -Origem (Get-LeadSourceFromText "Organico" @($orgRede, $orgStatus)) `
          -Plano "" `
          -Pipeline (Get-PipelineFromText @($orgStatus)) `
          -Observacoes $obs `
          -SourceFile (Split-Path $MarketingWorkbookPath -Leaf) `
          -SourceSheet "Vendas" `
          -SourceSection "Redes sociais organico" `
          -SourceRow $row)) | Out-Null
      } else {
        Add-SkippedRow $skippedRows "Lead organico sem nome ou contato valido." (Split-Path $MarketingWorkbookPath -Leaf) "Vendas" "Redes sociais organico" $row $orgNome $orgContato $orgStatus
      }
    }

    $paidData = Get-CellText $vendasSheet $row 9
    $paidQuantidade = Get-CellText $vendasSheet $row 10
    $paidContrato = Get-CellText $vendasSheet $row 11
    $paidNome = Get-CellText $vendasSheet $row 12
    $paidContato = Get-CellText $vendasSheet $row 13
    $paidRede = Get-CellText $vendasSheet $row 14
    $paidVendedor = Get-CellText $vendasSheet $row 15
    $paidStatus = Get-CellText $vendasSheet $row 16

    if (Has-MeaningfulText @($paidData, $paidQuantidade, $paidNome, $paidContato, $paidRede, $paidVendedor, $paidStatus)) {
      $cleanContact = Clean-Phone $paidContato
      if ((Normalize-Whitespace $paidNome) -and $cleanContact) {
        $obs = Build-Observation @($paidStatus)

        $allRows.Add((New-ImportRow `
          -Nome $paidNome `
          -Contato $paidContato `
          -Responsavel $paidVendedor `
          -Valor "" `
          -DataInicio $paidData `
          -RedeSocial $paidRede `
          -Origem (Get-LeadSourceFromText "Pago" @($paidRede, $paidStatus)) `
          -Plano "" `
          -NumeroContrato $paidContrato `
          -Pipeline (Get-PipelineFromText @($paidStatus)) `
          -Observacoes $obs `
          -SourceFile (Split-Path $MarketingWorkbookPath -Leaf) `
          -SourceSheet "Vendas" `
          -SourceSection "Redes sociais trafego pago" `
          -SourceRow $row)) | Out-Null
      } else {
        Add-SkippedRow $skippedRows "Lead pago sem nome ou contato valido." (Split-Path $MarketingWorkbookPath -Leaf) "Vendas" "Redes sociais trafego pago" $row $paidNome $paidContato $paidStatus
      }
    }

    $otherData = Get-CellText $vendasSheet $row 18
    $otherQuantidade = Get-CellText $vendasSheet $row 19
    $otherNome = Get-CellText $vendasSheet $row 20
    $otherContato = Get-CellText $vendasSheet $row 21
    $otherOrigem = Get-CellText $vendasSheet $row 22
    $otherStatus = Get-CellText $vendasSheet $row 23
    $otherResponsavel = Get-CellText $vendasSheet $row 24

    if (Has-MeaningfulText @($otherData, $otherQuantidade, $otherNome, $otherContato, $otherOrigem, $otherStatus, $otherResponsavel)) {
      $cleanContact = Clean-Phone $otherContato
      if ((Normalize-Whitespace $otherNome) -and $cleanContact) {
        $obs = Build-Observation @($otherStatus)

        $allRows.Add((New-ImportRow `
          -Nome $otherNome `
          -Contato $otherContato `
          -Responsavel $otherResponsavel `
          -Valor "" `
          -DataInicio $otherData `
          -RedeSocial $otherOrigem `
          -Origem (Get-LeadSourceFromText "Organico" @($otherOrigem, $otherStatus)) `
          -Plano "" `
          -Pipeline (Get-PipelineFromText @($otherStatus, $otherOrigem)) `
          -Observacoes $obs `
          -SourceFile (Split-Path $MarketingWorkbookPath -Leaf) `
          -SourceSheet "Vendas" `
          -SourceSection "Demais veiculos" `
          -SourceRow $row)) | Out-Null
      } else {
        Add-SkippedRow $skippedRows "Lead de veiculo externo sem nome ou contato valido." (Split-Path $MarketingWorkbookPath -Leaf) "Vendas" "Demais veiculos" $row $otherNome $otherContato $otherStatus
      }
    }
  }

  $marketingWorkbook.Close($false)
  $marketingWorkbook = $null

  $referralWorkbook = $excel.Workbooks.Open($ReferralWorkbookPath, 0, $true)
  $referralSheet = $referralWorkbook.Worksheets.Item(1)
  $referralRows = $referralSheet.UsedRange.Rows.Count
  $referralTimelineYear = 0
  $referralTimelineMonth = 0

  for ($row = 5; $row -le $referralRows; $row++) {
    $nome = Get-CellText $referralSheet $row 1
    $mesReferencia = Get-CellText $referralSheet $row 2
    $diaIndica = Get-CellText $referralSheet $row 3
    $contato = Get-CellText $referralSheet $row 4
    $setor = Get-CellText $referralSheet $row 5
    $colaborador = Get-CellText $referralSheet $row 6
    $lider = Get-CellText $referralSheet $row 7
    $vendedor = Get-CellText $referralSheet $row 8
    $status = Get-CellText $referralSheet $row 9
    $dataFechamento = Get-CellText $referralSheet $row 10
    $motivoNaoFechamento = Get-CellText $referralSheet $row 11
    $monthInfo = Parse-MonthReference $mesReferencia

    if ($monthInfo.Month -gt 0) {
      if ($monthInfo.Year -gt 0) {
        $referralTimelineYear = $monthInfo.Year
      } elseif ($referralTimelineYear -eq 0) {
        $referralTimelineYear = 2025
      } elseif ($referralTimelineMonth -gt 0 -and $monthInfo.Month -lt $referralTimelineMonth) {
        $referralTimelineYear++
      }

      if ($referralTimelineYear -gt 0) {
        $referralTimelineMonth = $monthInfo.Month
      }
    }

    if (-not (Has-MeaningfulText @($nome, $contato, $setor, $colaborador, $lider, $vendedor, $status, $motivoNaoFechamento))) {
      continue
    }

    $cleanContact = Clean-Phone $contato
    if (-not (Normalize-Whitespace $nome)) {
      Add-SkippedRow $skippedRows "Indicacao sem nome do lead." (Split-Path $ReferralWorkbookPath -Leaf) "PROGRAMA DE INDICACAO" "Indicacoes colaborador" $row $nome $contato $status
      continue
    }
    if (-not $cleanContact) {
      Add-SkippedRow $skippedRows "Indicacao sem contato valido." (Split-Path $ReferralWorkbookPath -Leaf) "PROGRAMA DE INDICACAO" "Indicacoes colaborador" $row $nome $contato $status
      continue
    }

    $obs = Build-Observation @($motivoNaoFechamento, $status)

    $startDate = Convert-DateToIso $diaIndica
    if (-not $startDate) {
      $startDate = Convert-DateToIso $dataFechamento
    }
    if (-not $startDate) {
      $startDate = Find-DateInText $motivoNaoFechamento $referralTimelineYear
    }
    if (-not $startDate) {
      $startDate = Find-DateInText $status $referralTimelineYear
    }
    if (-not $startDate -and $referralTimelineYear -gt 0 -and $referralTimelineMonth -gt 0) {
      $startDate = ("{0:D4}-{1:D2}-01" -f $referralTimelineYear, $referralTimelineMonth)
    }

    $allRows.Add((New-ImportRow `
      -Nome $nome `
      -Contato $contato `
      -Responsavel $vendedor `
      -Valor "" `
      -DataInicio $startDate `
      -RedeSocial $LabelIndicacaoColaborador `
      -Origem $LabelIndicacao `
      -IndicadoPor $colaborador `
      -SetorIndicado $setor `
      -Plano "" `
      -Pipeline (Get-PipelineFromText @($status, $motivoNaoFechamento)) `
      -Observacoes $obs `
      -SourceFile (Split-Path $ReferralWorkbookPath -Leaf) `
      -SourceSheet "PROGRAMA DE INDICACAO" `
      -SourceSection "Indicacoes colaborador" `
      -SourceRow $row)) | Out-Null
  }

  $referralWorkbook.Close($false)
  $referralWorkbook = $null
}
finally {
  if ($crmWorkbook -ne $null) { $crmWorkbook.Close($false) }
  if ($marketingWorkbook -ne $null) { $marketingWorkbook.Close($false) }
  if ($referralWorkbook -ne $null) { $referralWorkbook.Close($false) }
  if ($excel -ne $null) {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
}

$allRowsArray = [object[]]$allRows.ToArray()
$allRowsArray = @($allRowsArray | ForEach-Object {
  if (-not (Normalize-Whitespace $_.data_inicio)) {
    $_.data_inicio = Get-FallbackStartDate $_
  }
  Finalize-ImportRow $_
})

$allCsvColumns = @(
  "nome",
  "contato",
  "responsavel",
  "valor",
  "data_inicio",
  "rede_social",
  "origem",
  "indicado_por",
  "setor_indicado",
  "plano",
  "numero_contrato",
  "pipeline",
  "observacoes"
)

$auditColumns = @(
  "nome",
  "contato",
  "responsavel",
  "valor",
  "data_inicio",
  "rede_social",
  "origem",
  "indicado_por",
  "setor_indicado",
  "plano",
  "numero_contrato",
  "pipeline",
  "observacoes",
  "source_file",
  "source_sheet",
  "source_group",
  "source_row"
)

$deduplicatedRows = New-Object System.Collections.Generic.List[object]

foreach ($group in ($allRowsArray | Group-Object contato)) {
  $items = @($group.Group | Sort-Object `
    @{ Expression = { Get-PipelineRank $_.pipeline }; Descending = $true }, `
    @{ Expression = { $_.data_inicio }; Descending = $true }, `
    @{ Expression = { $_.source_row }; Descending = $true })

  $selected = $items[0]
  $otherItems = @($items | Select-Object -Skip 1)

  if ($otherItems.Count -gt 0) {
    $mergedNotes = @($items | ForEach-Object {
      $source = "$($_.source_file) > $($_.source_sheet) > $($_.source_group) > linha $($_.source_row)"
      $note = Normalize-Whitespace $_.observacoes
      if ($note) { "[$source] $note" } else { "[$source]" }
    } | Select-Object -Unique)

    $preferredOwner = @($items | ForEach-Object { Normalize-Whitespace $_.responsavel } | Where-Object { $_ }) | Select-Object -First 1
    $preferredSource = @($items | ForEach-Object { Normalize-Whitespace $_.rede_social } | Where-Object { $_ }) | Select-Object -First 1
    $preferredReferrer = @($items | ForEach-Object { Normalize-Whitespace $_.indicado_por } | Where-Object { $_ }) | Select-Object -First 1
    $preferredReferralSector = @($items | ForEach-Object { Normalize-Whitespace $_.setor_indicado } | Where-Object { $_ }) | Select-Object -First 1
    $preferredPlan = @($items | ForEach-Object { Normalize-Whitespace $_.plano } | Where-Object { $_ }) | Select-Object -First 1
    $preferredContractNumber = @($items | ForEach-Object { Normalize-Whitespace $_.numero_contrato } | Where-Object { $_ }) | Select-Object -First 1
    $selectedOriginKey = Normalize-LookupText $selected.origem

    $selected = [pscustomobject]@{
      nome         = $selected.nome
      contato      = $selected.contato
      responsavel  = if (Normalize-Whitespace $selected.responsavel) { $selected.responsavel } else { $preferredOwner }
      valor        = $selected.valor
      data_inicio  = if (Normalize-Whitespace $selected.data_inicio) { $selected.data_inicio } else { Get-FallbackStartDate $selected }
      rede_social  = if (Normalize-Whitespace $selected.rede_social) { $selected.rede_social } else { $preferredSource }
      origem       = $selected.origem
      indicado_por = if ($selectedOriginKey -eq "indicacao") { if (Normalize-Whitespace $selected.indicado_por) { $selected.indicado_por } else { $preferredReferrer } } else { "" }
      setor_indicado = if ($selectedOriginKey -eq "indicacao") { if (Normalize-Whitespace $selected.setor_indicado) { $selected.setor_indicado } else { $preferredReferralSector } } else { "" }
      plano        = if (Normalize-Whitespace $selected.plano) { $selected.plano } else { $preferredPlan }
      numero_contrato = if (Normalize-Whitespace $selected.numero_contrato) { $selected.numero_contrato } else { $preferredContractNumber }
      pipeline     = Get-PipelineFromText @($selected.pipeline, ($mergedNotes -join " || "))
      observacoes  = Normalize-Whitespace ($mergedNotes -join " || ")
      source_file  = $selected.source_file
      source_sheet = $selected.source_sheet
      source_group = $selected.source_group
      source_row   = $selected.source_row
    }

    foreach ($duplicate in $otherItems) {
      $duplicateRows.Add([pscustomobject]@{
        contato                  = $group.Name
        nome_selecionado         = $selected.nome
        pipeline_selecionado     = $selected.pipeline
        data_selecionada         = $selected.data_inicio
        origem_selecionada       = "$($selected.source_file) > $($selected.source_sheet) > $($selected.source_group) > linha $($selected.source_row)"
        nome_descartado          = $duplicate.nome
        pipeline_descartado      = $duplicate.pipeline
        data_descartada          = $duplicate.data_inicio
        origem_descartada        = "$($duplicate.source_file) > $($duplicate.source_sheet) > $($duplicate.source_group) > linha $($duplicate.source_row)"
      }) | Out-Null
    }
  }

  $deduplicatedRows.Add((Finalize-ImportRow $selected)) | Out-Null
}

$dateTag = Get-Date -Format "yyyyMMdd_HHmmss"

$allPath = Join-Path $OutputDir "crm_import_todos_$dateTag.csv"
$dedupPath = Join-Path $OutputDir "crm_import_deduplicado_$dateTag.csv"
$duplicatesPath = Join-Path $OutputDir "crm_import_duplicados_$dateTag.csv"
$skippedPath = Join-Path $OutputDir "crm_import_pulados_$dateTag.csv"
$auditPath = Join-Path $OutputDir "crm_import_auditoria_$dateTag.csv"
$summaryPath = Join-Path $OutputDir "crm_import_resumo_$dateTag.md"

Write-SemicolonCsv -Path $allPath -Rows $allRowsArray -Columns $allCsvColumns
Write-SemicolonCsv -Path $dedupPath -Rows ([object[]]$deduplicatedRows.ToArray()) -Columns $allCsvColumns
Write-SemicolonCsv -Path $auditPath -Rows $allRowsArray -Columns $auditColumns
Write-SemicolonCsv -Path $duplicatesPath -Rows ([object[]]$duplicateRows.ToArray()) -Columns @(
  "contato",
  "nome_selecionado",
  "pipeline_selecionado",
  "data_selecionada",
  "origem_selecionada",
  "nome_descartado",
  "pipeline_descartado",
  "data_descartada",
  "origem_descartada"
)
Write-SemicolonCsv -Path $skippedPath -Rows ([object[]]$skippedRows.ToArray()) -Columns @(
  "motivo",
  "source_file",
  "source_sheet",
  "source_group",
  "source_row",
  "nome",
  "contato",
  "status"
)

$summary = @(
  "# Resumo da conversao",
  "",
  "- Data de geracao: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
  "- Planilhas processadas: $(Split-Path $CrmWorkbookPath -Leaf), $(Split-Path $MarketingWorkbookPath -Leaf) e $(Split-Path $ReferralWorkbookPath -Leaf)",
  "- Linhas convertidas: $($allRowsArray.Count)",
  "- Linhas recomendadas para importacao apos deduplicacao: $($deduplicatedRows.Count)",
  "- Duplicidades removidas no arquivo recomendado: $($duplicateRows.Count)",
  "- Linhas puladas por falta de nome ou contato valido: $($skippedRows.Count)",
  "",
  "## Arquivos",
  "",
  "- Todos os registros convertidos: $(Split-Path $allPath -Leaf)",
  "- Arquivo recomendado para importar: $(Split-Path $dedupPath -Leaf)",
  "- Auditoria completa: $(Split-Path $auditPath -Leaf)",
  "- Duplicidades encontradas: $(Split-Path $duplicatesPath -Leaf)",
  "- Linhas puladas: $(Split-Path $skippedPath -Leaf)",
  "",
  "## Regras aplicadas",
  "",
  "- Telefones foram normalizados para apenas numeros e DDD 64 foi aplicado quando o numero tinha 8 ou 9 digitos.",
  "- Datas foram convertidas para o formato AAAA-MM-DD.",
  "- origem foi padronizada em Indicação, Evento Externo, Organico ou Pago para manter compatibilidade com o CRM atual.",
  "- pipeline foi inferido a partir do status original com os valores Novos Leads, Em Contato, Proposta, Fechado e Cancelado.",
  "- observacoes mantem a rastreabilidade da linha original e os detalhes da planilha."
)

$summaryEncoding = New-Object System.Text.UTF8Encoding($true)
[IO.File]::WriteAllText($summaryPath, ($summary -join "`r`n"), $summaryEncoding)

Write-Output "Arquivo recomendado: $dedupPath"
Write-Output "Linhas convertidas: $($allRowsArray.Count)"
Write-Output "Linhas recomendadas: $($deduplicatedRows.Count)"
Write-Output "Duplicidades removidas: $($duplicateRows.Count)"
Write-Output "Linhas puladas: $($skippedRows.Count)"
