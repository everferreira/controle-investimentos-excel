# Exporta prints reais da planilha via Excel COM (Windows) e faz dump de verificação
$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\conta\DIO\controle-investimentos-excel'
$wbPath = "$dir\controle-investimentos.xlsx"
$inv = [System.Globalization.CultureInfo]::InvariantCulture

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($wbPath)
$xl.CalculateFull()

$ws = $wb.Worksheets.Item('Simulador')

function V($row, $col) {
    $v = $ws.Cells.Item($row, $col).Value2
    if ($v -eq $null) { return '' }
    return $v.ToString($inv)
}
function Dump($tag) {
    Write-Output ("VAL|$tag|patrimonio=" + (V 12 2) + "|dividendos=" + (V 13 2))
    Write-Output ("VAL|$tag|proj2=" + (V 17 3) + "|proj5=" + (V 18 3) + "|proj10=" + (V 19 3) + "|proj20=" + (V 20 3) + "|proj30=" + (V 21 3))
    Write-Output ("VAL|$tag|pct_papel=" + (V 28 3) + "|pct_tijolo=" + (V 29 3) + "|pct_hib=" + (V 30 3) + "|pct_fofs=" + (V 31 3) + "|pct_dev=" + (V 32 3) + "|pct_hot=" + (V 33 3) + "|soma_pct=" + (V 34 3))
    Write-Output ("VAL|$tag|val_papel=" + (V 28 4) + "|val_tijolo=" + (V 29 4) + "|val_hib=" + (V 30 4) + "|val_fofs=" + (V 31 4) + "|val_dev=" + (V 32 4) + "|val_hot=" + (V 33 4) + "|total=" + (V 34 4))
}

# perfil 1: Conservador
$ws.Cells.Item(24, 2).Value2 = 'Conservador'
$xl.Calculate()
Dump 'conservador'
$ws.ExportAsFixedFormat(0, "$dir\prints\print_conservador.pdf")

# perfil 2: Agressivo
$ws.Cells.Item(24, 2).Value2 = 'Agressivo'
$xl.Calculate()
Dump 'agressivo'
$ws.ExportAsFixedFormat(0, "$dir\prints\print_agressivo.pdf")

# aba de apoio (independente do perfil)
$sup = $wb.Worksheets.Item('Perfis')
$sup.ExportAsFixedFormat(0, "$dir\prints\print_perfis.pdf")

# simulação com o aporte sugerido (30% do salário) — evidência extra
Write-Output ("VAL|sugestao10anos=" + (V 37 2) + "|sugestao_dividendos=" + (V 38 2))

# devolve o perfil padrão (Moderado), recalcula e SALVA (grava valores cacheados)
$ws.Cells.Item(24, 2).Value2 = 'Moderado'
$xl.Calculate()
$wb.Save()
$wb.Close($false)
$xl.Quit()
Write-Output 'EXPORT_OK'