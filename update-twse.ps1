<#
  update-twse.ps1
  抓取證交所（上市）與櫃買中心（上櫃）公開資料，計算籌碼 + 技術面，
  產出四頁儀表板資料，直接注入 taiwan-stock-dashboard.html（單一檔案）。

  關注清單 = 全部科技股（產業別 24–31），上市 + 上櫃約 900 檔。

  用法：  pwsh -File update-twse.ps1
  時機：  交易日收盤後約 18:00

  資料來源
    上市 TWSE：BFI82U（法人彙總）、T86（個股法人）、
               MI_INDEX（每日全市場收盤 → 兼作歷史股價、大盤、類股指數）、
               FMTQIK（加權指數歷史 → 相對強度）、t187ap03_L（公司產業別）
    上櫃 TPEx：insti/dailyTrade（個股法人）、tpex_mainboard_daily_close_quotes（收盤）、
               mopsfin_t187ap03_O（公司產業別）

  重要
    - 股價歷史改用「每日 MI_INDEX（單日全市場）」逐日快取，一天一次即涵蓋全部上市股，
      不再逐檔 STOCK_DAY（900 檔逐檔不可行）。過去日期不變 → 永久快取於 .cache\。
    - 上櫃個股無公開歷史股價端點，均線需逐日往後累積；首次執行僅一筆，技術面標資料不足。
#>

$ErrorActionPreference = 'Stop'
$HtmlFile = Join-Path $PSScriptRoot 'index.html'
$CacheDir = Join-Path $PSScriptRoot '.cache'
if (-not (Test-Path $CacheDir)) { New-Item -ItemType Directory -Path $CacheDir | Out-Null }

$InstDays  = 32   # 法人買賣超回溯天數（多抓 10 天，供「回推 10 天」時 20 日累計/連買仍完整）
$PriceDays = 72   # 股價回溯交易日（供 MA60；多抓 10 天，供回推 10 天仍算得出 60MA）
$TrendDays = 5
$AsOfDays  = 10   # 前端「資料日期」可回推的交易日數

# 官方產業別代碼 → 名稱（科技股宇宙 = 這 8 類）
$IndName = [ordered]@{
  '24'='半導體'; '25'='電腦及週邊'; '26'='光電'; '27'='通信網路'
  '28'='電子零組件'; '29'='電子通路'; '30'='資訊服務'; '31'='其他電子'
}

# 人工整理的「概念次產業」對照（散熱模組、AI 伺服器…官方分類沒有）。
# 這是產業地圖「概念族群」視圖；個股可同屬多概念。每群列較多代表公司。
# 代號會於下方以科技股清單驗證，不在清單者自動略過。
$Taxonomy = [ordered]@{
  'IC 設計'    = @{ parent='半導體與晶片'; codes=@('2454','3034','2379','3661','5269','6415','4966','3227','6104','6462','3443','4919','8016','3014','2401','6202','4952','5471') }
  '晶圓代工'   = @{ parent='半導體與晶片'; codes=@('2330','2303','6770','5347') }
  'IC 封測'    = @{ parent='半導體與晶片'; codes=@('3711','6239','2449','8150','3264','2441','6147','3374','2329','6257') }
  '矽智財 IP'  = @{ parent='半導體與晶片'; codes=@('6533','3529','6643','3035') }
  '記憶體'     = @{ parent='半導體與晶片'; codes=@('2344','2408','3006','4967','3260','2337','8299','2451','5289') }
  'AI 伺服器'  = @{ parent='AI 與伺服器'; codes=@('2317','2382','3231','6669','2356','2377','4938','2357','3706') }
  '散熱模組'   = @{ parent='AI 與伺服器'; codes=@('3017','6230','3653','3324','3338','3483','2421') }
  '機殼滑軌'   = @{ parent='AI 與伺服器'; codes=@('2059','3013','8210','6117') }
  '銅箔基板'   = @{ parent='AI 與伺服器'; codes=@('2383','6213','6274','6672') }
  'ABF 載板'   = @{ parent='AI 與伺服器'; codes=@('3037','8046','3189') }
  '5G 網通'    = @{ parent='網通與通訊'; codes=@('2345','3704','6285','2419','3596','5388','4906','2314') }
  '光通訊'     = @{ parent='網通與通訊'; codes=@('3081','4979','3363','4977','3234','3163','4908','3450') }
  '低軌衛星'   = @{ parent='網通與通訊'; codes=@('3491','2314','6197','6271') }
  '手機零組件' = @{ parent='網通與通訊'; codes=@('3008','2474','3406','2392','2439','3376','6205') }
  '筆電 NB'    = @{ parent='電腦與週邊'; codes=@('4938','2324','2382','2357','3231','2356') }
  '工業電腦'   = @{ parent='電腦與週邊'; codes=@('2395','3005','6414','8050','6166','3416','3022','6206') }
  '主機板'     = @{ parent='電腦與週邊'; codes=@('2357','2376','2377','3515') }
  '面板'       = @{ parent='光電與面板'; codes=@('2409','3481','6116','8069','6176','3673') }
  '光學鏡頭'   = @{ parent='光電與面板'; codes=@('3008','3406','3019','3362','6209','4976') }
  'Micro LED'  = @{ parent='光電與面板'; codes=@('3714','2393','4956') }
  'PCB'        = @{ parent='電子零組件'; codes=@('2313','6269','3037','2368','3044','5469','2316','6141') }
  '被動元件'   = @{ parent='電子零組件'; codes=@('2327','2492','2375','3026','2478','6173') }
  '連接器'     = @{ parent='電子零組件'; codes=@('3023','3665','3526','3605','3003','6290','3501') }
}
# $SectorsOf 於科技股清單建立後再驗證產生（見下）

# ---------------------------------------------------------------- 工具
function Get-Json { param([string]$Url,[int]$Retry=3)
  for ($i=1;$i -le $Retry;$i++){
    try { return Invoke-RestMethod -Uri $Url -TimeoutSec 90 -UserAgent 'Mozilla/5.0 (twse-dashboard)' }
    catch { if ($i -eq $Retry){ throw "取得 $Url 失敗：$($_.Exception.Message)" }; Start-Sleep -Seconds (3*$i) }
  }
}
# 對 307 節流退避，失敗回 $null（不中止）
function Get-JsonSoft { param([string]$Url)
  for ($a=1;$a -le 4;$a++){
    try { return Invoke-RestMethod -Uri $Url -TimeoutSec 90 -UserAgent 'Mozilla/5.0 (twse-dashboard)' -MaximumRedirection 0 -ErrorAction Stop }
    catch { Start-Sleep -Milliseconds (900*$a) }
  }
  return $null
}
function N { param($T) $s=("$T").Replace(',','').Replace('+','').Trim()
  if ($s -in @('','--','-','N/A','X','null')){ return 0 }; $n=0.0; if ([double]::TryParse($s,[ref]$n)){ return $n }; return 0 }
function Mark { param($T) ([regex]::Replace("$T",'<[^>]*>','')).Trim() }
function RocToIso { param($R) $p=("$R").Trim() -split '/'; if ($p.Count -ne 3){ return $null }
  '{0:0000}-{1:00}-{2:00}' -f ([int]$p[0]+1911),[int]$p[1],[int]$p[2] }

# ---------------------------------------------------------------- 科技股宇宙
Write-Host '建立科技股清單（產業別 24–31）...' -ForegroundColor Cyan
$L = Get-Json 'https://openapi.twse.com.tw/v1/opendata/t187ap03_L'
$O = Get-Json 'https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O'
$uni = @{}
foreach ($p in $L) { $ind=$p.'產業別'; if ($IndName.Contains($ind)) {
  $uni[$p.'公司代號'] = @{ name=$p.'公司簡稱'; market='TWSE'; indName=$IndName[$ind] } } }
foreach ($p in $O) { $c=("$($p.SecuritiesCompanyCode)").Trim(); $ind=$p.SecuritiesIndustryCode
  if ($IndName.Contains($ind) -and -not $uni.ContainsKey($c)) {
    $uni[$c] = @{ name=$p.CompanyAbbreviation; market='TPEx'; indName=$IndName[$ind] } } }
$AllCodes = @($uni.Keys)
Write-Host "  科技股 $($AllCodes.Count) 檔（上市 $((@($uni.Values|Where-Object{$_.market -eq 'TWSE'})).Count) + 上櫃 $((@($uni.Values|Where-Object{$_.market -eq 'TPEx'})).Count)）" -ForegroundColor Green

# 驗證概念族群代號都在科技股清單內，並建立個股→概念的反查
foreach ($g in @($Taxonomy.Keys)) {
  $valid=@($Taxonomy[$g].codes | Where-Object { $uni.ContainsKey($_) })
  $bad=@($Taxonomy[$g].codes | Where-Object { -not $uni.ContainsKey($_) })
  if ($bad.Count){ Write-Host "  概念族群 [$g] 略過非科技股代號：$($bad -join ',')" -ForegroundColor DarkYellow }
  $Taxonomy[$g].codes=$valid
}
$SectorsOf = @{}
foreach ($g in $Taxonomy.GetEnumerator()) {
  foreach ($c in $g.Value.codes) { if (-not $SectorsOf.ContainsKey($c)) { $SectorsOf[$c]=@() }; $SectorsOf[$c]+=$g.Key }
}

# ---------------------------------------------------------------- 交易日（法人）
Write-Host "尋找最近 $InstDays 個交易日 ..." -ForegroundColor Cyan
$days=@(); $probe=(Get-Date).Date; $guard=0
while ($days.Count -lt $InstDays -and $guard -lt 45) {
  $guard++
  if ($probe.DayOfWeek -in @('Saturday','Sunday')){ $probe=$probe.AddDays(-1); continue }
  $ymd=$probe.ToString('yyyyMMdd')
  $r = Get-Json "https://www.twse.com.tw/rwd/zh/fund/BFI82U?dayDate=$ymd&type=day&response=json"
  if ($r.stat -eq 'OK' -and $r.data) {
    $v=@{}; foreach ($d in $r.data){ $v[$d[0]]=(N $d[3]) }
    $days += [pscustomobject]@{ date=$probe.ToString('yyyy-MM-dd'); ymd=$ymd
      slash=$probe.ToString('yyyy/MM/dd'); label=$probe.ToString('MM/dd')
      foreign=[math]::Round((($v['外資及陸資(不含外資自營商)']+$v['外資自營商'])/1e8),1)
      trust=[math]::Round(($v['投信']/1e8),1)
      dealer=[math]::Round((($v['自營商(自行買賣)']+$v['自營商(避險)'])/1e8),1) }
  }
  $probe=$probe.AddDays(-1); Start-Sleep -Milliseconds 400
}
if ($days.Count -eq 0){ throw '找不到交易日資料。' }
[array]::Reverse($days)
# 優先用當日；若當日上市盤後收盤尚未備妥（還在盤中／未публ），退回前一交易日。
# 法人（BFI82U）與收盤（MI_INDEX）皆為盤後才發布，兩者一致才採用，避免混日。
$miLatest=$null
while ($days.Count -ge 1) {
  $cand=$days[-1]
  $mi=Get-JsonSoft "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date=$($cand.ymd)&type=ALLBUT0999&response=json"
  if ($mi -and $mi.stat -eq 'OK' -and "$($mi.date)" -eq $cand.ymd) { $miLatest=$mi; break }
  if ($days.Count -eq 1){ throw '找不到已備妥的交易日收盤資料。' }
  Write-Host "  $($cand.date) 上市盤後資料尚未備妥，改用前一交易日" -ForegroundColor Yellow
  $days=@($days[0..($days.Count-2)])
}
$latest=$days[-1]; $dateList=@($days|ForEach-Object{$_.date})
$isToday = ($latest.date -eq (Get-Date).ToString('yyyy-MM-dd'))
Write-Host "  採用交易日 $($latest.date)（$(if($isToday){'當日'}else{'最近有資料日'})），$($days.Count) 日" -ForegroundColor Green

# ---------------------------------------------------------------- 多日籌碼（個股，含每日快取）
Write-Host '抓取多日三大法人買賣超（個股）...' -ForegroundColor Cyan
$hist=@{}   # code -> date -> @{f,t,d,net}
foreach ($d in $days) {
  $cache = Join-Path $CacheDir "inst_$($d.ymd).json"
  $rows = $null
  if ($d.date -ne $latest.date -and (Test-Path $cache)) {
    $rows = Get-Content $cache -Raw | ConvertFrom-Json
  } else {
    $rows = @()
    $t = Get-Json "https://www.twse.com.tw/rwd/zh/fund/T86?date=$($d.ymd)&selectType=ALL&response=json"
    if ($t.stat -eq 'OK') { foreach ($row in $t.data) { $c=("$($row[0])").Trim(); if (-not $uni.ContainsKey($c)){ continue }
      $rows += [pscustomobject]@{ code=$c; f=[math]::Round(((N $row[4])+(N $row[7]))/1000); t=[math]::Round((N $row[10])/1000); d=[math]::Round((N $row[11])/1000); net=[math]::Round((N $row[18])/1000) } } }
    Start-Sleep -Milliseconds 450
    $o = Get-Json "https://www.tpex.org.tw/www/zh-tw/insti/dailyTrade?type=Daily&sect=EW&date=$($d.slash)&id=&response=json"
    if ($o.tables -and $o.tables[0].data) { foreach ($row in $o.tables[0].data) { $c=("$($row[0])").Trim(); if (-not $uni.ContainsKey($c)){ continue }
      $rows += [pscustomobject]@{ code=$c; f=[math]::Round((N $row[10])/1000); t=[math]::Round((N $row[13])/1000); d=[math]::Round((N $row[22])/1000); net=[math]::Round((N $row[23])/1000) } } }
    if ($d.date -ne $latest.date) { $rows | ConvertTo-Json -Compress | Set-Content $cache -Encoding UTF8 }
    Start-Sleep -Milliseconds 450
  }
  foreach ($x in $rows) { if (-not $hist.ContainsKey($x.code)){ $hist[$x.code]=@{} }
    $hist[$x.code][$d.date]=@{ f=[double]$x.f; t=[double]$x.t; d=[double]$x.d; net=[double]$x.net } }
  Write-Host ("  {0}  {1,4} 檔" -f $d.label, $rows.Count)
}

# ---------------------------------------------------------------- 股價歷史（每日 MI_INDEX 全市場，含快取）
Write-Host '抓取每日全市場收盤（算均線，含快取）...' -ForegroundColor Cyan
# 交易日軸：以 FMTQIK（加權指數）日期為準
$idxHist=@{}
$mp=Get-Date -Day 1; $ymList=@()
for ($i=0;$i -lt 5;$i++){ $ymList+=$mp.ToString('yyyyMM'); $mp=$mp.AddMonths(-1) }
[array]::Reverse($ymList); $curMonth=(Get-Date).ToString('yyyyMM')
foreach ($ym in $ymList) {
  $cache=Join-Path $CacheDir "taiex_$ym.json"; $rows=$null
  if ($ym -ne $curMonth -and (Test-Path $cache)){ $rows=Get-Content $cache -Raw|ConvertFrom-Json }
  else { $r=Get-JsonSoft "https://www.twse.com.tw/rwd/zh/afterTrading/FMTQIK?date=${ym}01&response=json"
    $rows=@(); if ($r.stat -eq 'OK'){ foreach ($row in $r.data){ $iso=RocToIso $row[0]; if ($iso){ $rows+=[pscustomobject]@{ date=$iso; close=(N $row[4]) } } } }
    if ($ym -ne $curMonth -and $rows.Count){ $rows|ConvertTo-Json -Compress|Set-Content $cache -Encoding UTF8 }
    Start-Sleep -Milliseconds 400 }
  foreach ($x in $rows){ $idxHist[$x.date]=$x.close }
}
$axis=@($idxHist.Keys | Sort-Object | Select-Object -Last $PriceDays)

# 逐日 MI_INDEX：回傳 code->@{close,vol}，過去日永久快取
function Get-DayCloses { param([string]$Iso)
  $ymd = $Iso -replace '-',''
  $cache = Join-Path $CacheDir "mi_$ymd.json"
  if ($Iso -ne $latest.date -and (Test-Path $cache)) {
    $o = Get-Content $cache -Raw | ConvertFrom-Json
    $h=@{}; foreach ($p in $o.PSObject.Properties){ $h[$p.Name]=@{ close=[double]$p.Value[0]; vol=[double]$p.Value[1] } }; return $h
  }
  $mi = Get-JsonSoft "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date=$ymd&type=ALLBUT0999&response=json"
  $map=@{}
  if ($mi -and $mi.stat -eq 'OK') {
    $pt = $mi.tables | Where-Object { $_.fields -contains '收盤價' } | Select-Object -First 1
    if ($pt) { foreach ($r in $pt.data){ $c=("$($r[0])").Trim(); if (-not $uni.ContainsKey($c)){ continue }
      $close=N $r[8]; if ($close -le 0){ continue }; $map[$c]=@{ close=$close; vol=[math]::Round((N $r[2])/1000) } } }
  }
  if ($Iso -ne $latest.date -and $map.Count) {
    $out=[ordered]@{}; foreach ($k in $map.Keys){ $out[$k]=@($map[$k].close,$map[$k].vol) }
    [pscustomobject]$out | ConvertTo-Json -Compress | Set-Content $cache -Encoding UTF8
  }
  Start-Sleep -Milliseconds 450
  return $map
}
# 逐日建立股價序列（上市）
$priceSeries=@{}   # code -> date -> @{close,vol}
$di=0
foreach ($iso in $axis) {
  $di++
  $dayMap = Get-DayCloses $iso
  foreach ($c in $dayMap.Keys) { if (-not $priceSeries.ContainsKey($c)){ $priceSeries[$c]=@{} }
    $priceSeries[$c][$iso]=$dayMap[$c] }
  if ($di % 10 -eq 0){ Write-Host "  ...股價 $di/$($axis.Count) 天" }
}

# 最新收盤／量／漲跌（上市，取自已於交易日確認階段取得的 $miLatest）
Write-Host '解析最新收盤與大盤 ...' -ForegroundColor Cyan
$quoteBy=@{}
$pt = $miLatest.tables | Where-Object { $_.fields -contains '收盤價' } | Select-Object -First 1
foreach ($r in $pt.data) { $c=("$($r[0])").Trim(); if (-not $uni.ContainsKey($c)){ continue }
  $close=N $r[8]; if ($close -le 0){ continue }
  $m=Mark $r[9]; $diff=N $r[10]; $chg= if ($m -eq '-'){ -$diff } elseif ($m -eq '+'){ $diff } else { 0 }; $prev=$close-$chg
  $quoteBy[$c]=@{ name=$uni[$c].name; market='TWSE'; close=$close
    pct= if ($prev -ne 0){ [math]::Round($chg/$prev*100,2) } else { 0 }
    amount=N $r[4]; volLots=[math]::Round((N $r[2])/1000) } }

# 上櫃：最新報價 + 往後累積股價快取
$tq = Get-Json 'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes'
$rocLatest = ([int]$latest.date.Substring(0,4)-1911).ToString()+$latest.date.Substring(5,2)+$latest.date.Substring(8,2)
$tpexCacheFile = Join-Path $CacheDir 'tpex_close.json'
$tpexClose=@{}
if (Test-Path $tpexCacheFile){ $obj=Get-Content $tpexCacheFile -Raw|ConvertFrom-Json
  foreach ($p in $obj.PSObject.Properties){ $tpexClose[$p.Name]=@{}; foreach ($q in $p.Value.PSObject.Properties){ $tpexClose[$p.Name][$q.Name]=$q.Value } } }
if ((($tq|Select-Object -First 1).Date) -eq $rocLatest) {
  foreach ($r in $tq) { $c=("$($r.SecuritiesCompanyCode)").Trim(); if (-not $uni.ContainsKey($c)){ continue }
    $close=N $r.Close; if ($close -le 0){ continue }
    $chg=0.0; if (("$($r.Change)").Trim() -match '^[+-]?[\d,.]+$'){ $chg=[double](("$($r.Change)") -replace ',','') }; $prev=$close-$chg
    $quoteBy[$c]=@{ name=$uni[$c].name; market='TPEx'; close=$close
      pct= if ($prev -ne 0){ [math]::Round($chg/$prev*100,2) } else { 0 }
      amount=N $r.TransactionAmount; volLots=[math]::Round((N $r.TradingShares)/1000) }
    if (-not $tpexClose.ContainsKey($c)){ $tpexClose[$c]=@{} }
    $tpexClose[$c][$latest.date]=$close }
} else { Write-Host "  警告：上櫃報價日期與 $rocLatest 不符。" -ForegroundColor Yellow }
foreach ($c in $tpexClose.Keys) { if (-not $priceSeries.ContainsKey($c)){ $priceSeries[$c]=@{} }
  foreach ($dt in $tpexClose[$c].Keys){ $priceSeries[$c][$dt]=@{ close=[double]$tpexClose[$c][$dt]; vol=$null } } }
# 存回上櫃累積
($tpexClose.Keys | ForEach-Object { [pscustomobject]@{ code=$_; d=$tpexClose[$_] } } |
  ForEach-Object -Begin { $h=[ordered]@{} } -Process { $h[$_.code]=$_.d } -End { [pscustomobject]$h }) |
  ConvertTo-Json -Depth 5 -Compress | Set-Content $tpexCacheFile -Encoding UTF8

# 大盤統計 + 類股指數（取自最新 MI_INDEX）
$idxTable = $miLatest.tables | Where-Object { $_.fields -contains '收盤指數' } | Select-Object -First 1
$twii = $idxTable.data | Where-Object { $_[0] -eq '發行量加權股價指數' } | Select-Object -First 1
$twiiMark = Mark $twii[2]
$sumTable = $miLatest.tables | Where-Object { $_.title -match '大盤統計資訊' } | Select-Object -First 1
$sumRow = $sumTable.data | Where-Object { $_[0] -match '^1\.一般股票' } | Select-Object -First 1
$udTable = $miLatest.tables | Where-Object { $_.title -match '漲跌證券數' } | Select-Object -First 1
function UD { param($La) (($udTable.data | Where-Object { $_[0] -match $La } | Select-Object -First 1)[1]) }
$market=[ordered]@{ index=N $twii[1]
  pct= if ($twiiMark -eq '-'){ -(N $twii[4]) } else { N $twii[4] }
  points= if ($twiiMark -eq '-'){ -(N $twii[3]) } else { N $twii[3] }
  amount=[math]::Round((N $sumRow[1])/1e8,0); up="$(UD '上漲')"; down="$(UD '下跌')"; flat="$(UD '持平')" }

# ---------------------------------------------------------------- 均線工具
$idxAxisClose=@($axis | ForEach-Object { if ($idxHist.ContainsKey($_)){ [double]$idxHist[$_] } else { $null } })
function MA { param($Arr,$Nn) $vals=@($Arr|Where-Object{$_ -ne $null}); if ($vals.Count -lt $Nn){ return $null }
  [math]::Round((($vals[-$Nn..-1]|Measure-Object -Sum).Sum/$Nn),2) }
function Ret { param($Arr,$Nn) $vals=@($Arr|Where-Object{$_ -ne $null}); if ($vals.Count -le $Nn){ return $null }
  $a=$vals[-1]; $b=$vals[-($Nn+1)]; if ($b -eq 0){ return $null }; [math]::Round(($a-$b)/$b*100,2) }
$idxRet20 = Ret $idxAxisClose 20

# ---------------------------------------------------------------- 逐檔綜合
Write-Host '計算技術面 × 籌碼面 ...' -ForegroundColor Cyan
function SumLast { param($Arr,$Nn) [int]((@($Arr|Select-Object -Last $Nn)|Measure-Object -Sum).Sum) }
$stocks=@(); $noquote=0
foreach ($c in ($AllCodes | Sort-Object -Unique)) {
  if (-not $quoteBy.ContainsKey($c)){ $noquote++; continue }
  $q=$quoteBy[$c]
  $ps= if ($priceSeries.ContainsKey($c)){ $priceSeries[$c] } else { @{} }
  $closes=@($axis | ForEach-Object { if ($ps.ContainsKey($_)){ [double]$ps[$_].close } else { $null } })
  $vols=@($axis | ForEach-Object { if ($ps.ContainsKey($_) -and $ps[$_].vol -ne $null){ [double]$ps[$_].vol } else { $null } })
  $ma5=MA $closes 5; $ma20=MA $closes 20; $ma60=MA $closes 60
  $hasTech=($ma5 -ne $null -and $ma20 -ne $null -and $ma60 -ne $null)
  $close=$q.close
  $r20=Ret $closes 20; $rs20= if ($r20 -ne $null -and $idxRet20 -ne $null){ [math]::Round($r20-$idxRet20,2) } else { $null }
  $multiHead=($hasTech -and $close -gt $ma5 -and $ma5 -gt $ma20 -and $ma20 -gt $ma60)
  $aboveMA20=($ma20 -ne $null -and $close -gt $ma20)
  $techGrade= if (-not $hasTech){ 'na' } elseif ($multiHead){ '強' } elseif ($aboveMA20){ '中' } else { '弱' }
  $distMA20= if ($ma20 -ne $null -and $ma20 -ne 0){ [math]::Round(($close-$ma20)/$ma20*100,1) } else { $null }
  $volAvg20=MA $vols 20; $withVol=($volAvg20 -ne $null -and $q.volLots -gt $volAvg20)

  $ft=@($dateList | ForEach-Object { if ($hist.ContainsKey($c) -and $hist[$c].ContainsKey($_)){ [double]($hist[$c][$_].f+$hist[$c][$_].t) } else { 0 } })
  $fArr=@($dateList | ForEach-Object { if ($hist.ContainsKey($c) -and $hist[$c].ContainsKey($_)){ [int]$hist[$c][$_].f } else { 0 } })
  $tArr=@($dateList | ForEach-Object { if ($hist.ContainsKey($c) -and $hist[$c].ContainsKey($_)){ [int]$hist[$c][$_].t } else { 0 } })
  $dToday= if ($hist.ContainsKey($c) -and $hist[$c].ContainsKey($latest.date)){ [int]$hist[$c][$latest.date].d } else { 0 }
  $ftStreak=0; $ftSign=[math]::Sign($ft[-1])
  if ($ftSign -ne 0){ for ($i=$ft.Count-1;$i -ge 0;$i--){ if ([math]::Sign($ft[$i]) -eq $ftSign){ $ftStreak++ } else { break } } }
  $acc5=SumLast $ft 5; $acc10=SumLast $ft 10; $acc20=SumLast $ft 20
  $todayFt=[int]$ft[-1]
  $ratio= if ($q.volLots -gt 0){ [math]::Round($todayFt/$q.volLots*100,1) } else { 0 }
  $priceSync=($todayFt -gt 0 -and $q.pct -gt 0)
  $chipStrong=($ftStreak -ge 3 -and $ftSign -gt 0)
  $chipGrade= if ($chipStrong){ '強' } elseif ($acc5 -gt 0){ '中' } else { '弱' }
  $techStrongAxis=($techGrade -in @('強','中'))
  $quad= if (-not $hasTech){ 'na' } elseif ($techStrongAxis -and $chipStrong){ 'main' }
    elseif ($techStrongAxis -and -not $chipStrong){ 'retail' }
    elseif (-not $techStrongAxis -and $chipStrong){ 'stealth' } else { 'exclude' }
  $score=0
  if ($multiHead){ $score+=25 }; if ($aboveMA20){ $score+=10 }
  if ($rs20 -ne $null -and $rs20 -gt 0){ $score+=20 }
  $score+=[math]::Round([math]::Min($ftStreak,5)/5*20)
  if ($acc5 -gt 0){ $score+=10 }; if ($withVol){ $score+=10 }; if ($priceSync){ $score+=5 }
  if (-not $hasTech){ $score=$null }

  # closes / vols 只在有足夠歷史時注入（省體積；上櫃多為空）。供前端回推日期重算。
  $hasSeries = (@($closes|Where-Object{$_ -ne $null}).Count -ge 20)
  $closesOut = if ($hasSeries) { @($closes | ForEach-Object { if ($_ -eq $null){ $null } else { [math]::Round($_,2) } }) } else { @() }
  $volsOut   = if ($hasSeries) { @($vols   | ForEach-Object { if ($_ -eq $null){ $null } else { [int]$_ } }) } else { @() }

  $stocks += [pscustomobject]@{
    code=$c; name=$q.name; market=$q.market; ind=$uni[$c].indName; sectors=@($SectorsOf[$c])
    close=$close; pct=$q.pct; volLots=$q.volLots; amount=[math]::Round($q.amount/1e8,2)
    ma5=$ma5; ma20=$ma20; ma60=$ma60; distMA20=$distMA20; rs20=$rs20
    withVol=$withVol; multiHead=$multiHead; aboveMA20=$aboveMA20; techGrade=$techGrade
    foreign=[int]$fArr[-1]; trust=[int]$tArr[-1]; dealer=$dToday
    ftToday=$todayFt; ftStreak=$ftStreak; ftSign=$ftSign; ftCapped=($ftStreak -ge $ft.Count -and $ftSign -ne 0)
    acc5=$acc5; acc10=$acc10; acc20=$acc20; ratio=$ratio; priceSync=$priceSync
    chipGrade=$chipGrade; quad=$quad; score=$score
    netAmt=[math]::Round($todayFt*1000*$close/1e8,2)
    closes=$closesOut; vols=$volsOut
    fSeries=@(for ($i=0;$i -lt $fArr.Count;$i++){ [int]$fArr[$i] })
    tSeries=@(for ($i=0;$i -lt $tArr.Count;$i++){ [int]$tArr[$i] })
  }
}
Write-Host "  完成 $($stocks.Count) 檔（$noquote 檔當日無報價略過）" -ForegroundColor Green

# ---------------------------------------------------------------- 族群彙總（共用）
function Build-Group { param($Name,$Parent,$Members)
  if ($Members.Count -eq 0){ return $null }
  $amtSum=($Members|Measure-Object amount -Sum).Sum
  $wPct= if ($amtSum -gt 0){ [math]::Round((($Members|ForEach-Object{$_.pct*$_.amount}|Measure-Object -Sum).Sum/$amtSum),2) } else { [math]::Round((($Members|Measure-Object pct -Average).Average),2) }
  $daily=@(); for ($i=0;$i -lt $dateList.Count;$i++){ $s=0; foreach ($m in $Members){ $s+=($m.fSeries[$i]+$m.tSeries[$i]) }; $daily+=$s }
  $streak=0; $ssign=[math]::Sign($daily[-1])
  if ($ssign -ne 0){ for ($i=$daily.Count-1;$i -ge 0;$i--){ if ([math]::Sign($daily[$i]) -eq $ssign){ $streak++ } else { break } } }
  $lead=$Members|Sort-Object pct -Descending|Select-Object -First 1
  [pscustomobject]@{ name=$Name; parent=$Parent; count=$Members.Count; pct=$wPct
    amount=[math]::Round($amtSum,1)
    foreign=[int](($Members|Measure-Object foreign -Sum).Sum); trust=[int](($Members|Measure-Object trust -Sum).Sum)
    dealer=[int](($Members|Measure-Object dealer -Sum).Sum); ftNet=[int](($Members|Measure-Object ftToday -Sum).Sum)
    netAmt=[math]::Round((($Members|Measure-Object netAmt -Sum).Sum),1)
    streak=$streak; streakSign=$ssign; streakCapped=($streak -ge $daily.Count -and $ssign -ne 0)
    lead=[ordered]@{ code=$lead.code; name=$lead.name; pct=$lead.pct }
    members=@($Members.code) }
}
# 概念族群（人工 23）
$groups=@()
foreach ($g in $Taxonomy.GetEnumerator()) {
  $mem=@($stocks|Where-Object{ $_.sectors -contains $g.Key })
  $grp=Build-Group $g.Key $g.Value.parent $mem; if ($grp){ $groups+=$grp }
}
$groups=@($groups|Sort-Object{[math]::Abs($_.netAmt)} -Descending)
# 產業別（官方 8，涵蓋全部）
$industryGroups=@()
foreach ($iname in $IndName.Values) {
  $mem=@($stocks|Where-Object{ $_.ind -eq $iname })
  $grp=Build-Group $iname '產業別' $mem; if ($grp){ $industryGroups+=$grp }
}
$industryGroups=@($industryGroups|Sort-Object{[math]::Abs($_.netAmt)} -Descending)

$parents=@()
foreach ($p in ($Taxonomy.Values.parent|Select-Object -Unique)) {
  $parents+=[pscustomobject]@{ name=$p; groups=@($groups|Where-Object{$_.parent -eq $p}|ForEach-Object{$_.name}) }
}

# ---------------------------------------------------------------- 三大法人（可查日期）
$instByDate=@()
foreach ($d in $days) {
  $rf=@()
  foreach ($s in $stocks) { if (-not ($hist.ContainsKey($s.code) -and $hist[$s.code].ContainsKey($d.date))){ continue }
    $h=$hist[$s.code][$d.date]; $rf+=[pscustomobject]@{ code=$s.code; name=$s.name; f=[int]$h.f; t=[int]$h.t } }
  $instByDate += [ordered]@{ date=$d.date; foreign=$d.foreign; trust=$d.trust; dealer=$d.dealer
    foreignTop=@($rf|Sort-Object f -Descending|Select-Object -First 8|ForEach-Object{[ordered]@{code=$_.code;name=$_.name;v=$_.f}})
    foreignBot=@($rf|Sort-Object f|Select-Object -First 8|ForEach-Object{[ordered]@{code=$_.code;name=$_.name;v=$_.f}})
    trustTop=@($rf|Sort-Object t -Descending|Select-Object -First 8|ForEach-Object{[ordered]@{code=$_.code;name=$_.name;v=$_.t}})
    trustBot=@($rf|Sort-Object t|Select-Object -First 8|ForEach-Object{[ordered]@{code=$_.code;name=$_.name;v=$_.t}}) }
}

# ---------------------------------------------------------------- 法人總覽趨勢
$trend=@($days|Select-Object -Last $TrendDays)
$institutions=@(
  @{key='foreign';name='外資';en='FOREIGN'}; @{key='trust';name='投信';en='TRUST'}; @{key='dealer';name='自營商';en='DEALER'}
) | ForEach-Object { $k=$_.key
  $series=@($trend|ForEach-Object{[double]$_.$k}); $full=@($days|ForEach-Object{[double]$_.$k})
  $nn=0; $sign=[math]::Sign($full[-1]); if ($sign -ne 0){ for ($i=$full.Count-1;$i -ge 0;$i--){ if ([math]::Sign($full[$i]) -eq $sign){ $nn++ } else { break } } }
  [ordered]@{ key=$k; name=$_.name; en=$_.en; today=$series[-1]
    sum5=[math]::Round((($series|Measure-Object -Sum).Sum),1); series=$series
    streakDays=$nn; streakSign=$sign; streakCapped=($nn -ge $full.Count -and $sign -ne 0) } }

# ---------------------------------------------------------------- 輸出
$idxClosesOut=@($idxAxisClose | ForEach-Object { if ($_ -eq $null){ $null } else { [math]::Round($_,2) } })
$payload=[ordered]@{
  generatedAt=(Get-Date).ToString('yyyy-MM-dd HH:mm'); tradeDate=$latest.date
  universe=$stocks.Count; asOfDays=$AsOfDays
  priceAxis=$axis; idxRet20=$idxRet20; idxCloses=$idxClosesOut
  trendLabels=@($trend|ForEach-Object{$_.label}); histLabels=@($days|ForEach-Object{$_.label}); histDates=$dateList
  market=$market; institutions=$institutions; instByDate=$instByDate
  parents=$parents; groups=$groups; industryGroups=$industryGroups; stocks=$stocks
}
$json=$payload | ConvertTo-Json -Depth 12 -Compress
if (-not (Test-Path $HtmlFile)){ throw "找不到 $HtmlFile" }
$html=Get-Content $HtmlFile -Raw -Encoding UTF8
$s='/* DATA-START */'; $e='/* DATA-END */'
if ($html -notmatch [regex]::Escape($s)){ throw "HTML 缺少 $s 標記" }
$block="$s`nwindow.TWSE_DATA = $json;`n$e"
$html=[regex]::Replace($html,([regex]::Escape($s)+'.*?'+[regex]::Escape($e)),{ $block },'Singleline')
Set-Content -Path $HtmlFile -Value $html -Encoding UTF8 -NoNewline

$strong=@($stocks|Where-Object{$_.quad -eq 'main'}).Count
$tech=@($stocks|Where-Object{$_.techGrade -ne 'na'}).Count
Write-Host ''
Write-Host "已更新 $HtmlFile" -ForegroundColor Green
Write-Host "$($latest.date) · 大盤 $($market.index) ($($market.pct)%) · 科技股 $($stocks.Count) 檔（技術面可算 $tech）· 主流強勢股 $strong 檔" -ForegroundColor Green
Write-Host "JSON 約 $([math]::Round($json.Length/1MB,2)) MB" -ForegroundColor DarkGray
