Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Normalize-Text([string]$Text) {
    if ($null -eq $Text) { return '' }
    return $Text.Normalize([Text.NormalizationForm]::FormC)
}

function Compact-Text([string]$Text) {
    $n = Normalize-Text $Text
    return [regex]::Replace($n, '[\s_\-\*·\.\(\)\[\]]', '')
}

function Split-Names([string]$Text) {
    return @($Text -split '[,，;；\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Pick-Folder([string]$Initial) {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '폴더를 선택하세요.'
    if (Test-Path -LiteralPath $Initial) { $dialog.SelectedPath = $Initial }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
    return $null
}

function Safe-Name([string]$Name) {
    $result = $Name
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $result = $result.Replace([string]$c, '_') }
    return $result.Trim()
}

function Is-MonthMatch([string]$Name, [string]$MonthRange) {
    if ([string]::IsNullOrWhiteSpace($MonthRange)) { return $false }
    $digits = @([regex]::Matches($MonthRange, '\d+') | ForEach-Object { $_.Value })
    if ($digits.Count -lt 1) { return $false }
    $normalizedName = Normalize-Text $Name
    if ($digits.Count -eq 1) {
        $month = [regex]::Escape($digits[0])
        $pattern = '(?<!\d)(?:\(\s*' + $month + '\s*\)\s*월|\(\s*' + $month + '\s*월\s*\)|' + $month + '\s*월)(?!\d)'
        return $normalizedName -match $pattern
    }
    $pattern = '(?<!\d)' + [regex]::Escape($digits[0]) + '\s*(?:월)?\s*[-~]\s*' + [regex]::Escape($digits[1]) + '\s*(?:월)?(?!\d)'
    return $normalizedName -match $pattern
}

function Get-CohortNumbers([string]$Name) {
    return @([regex]::Matches((Normalize-Text $Name), '(?<!\d)(\d+)\s*기(?!\d)') | ForEach-Object { [int]$_.Groups[1].Value })
}

function Get-CohortNumber([string]$Value) {
    $match = [regex]::Match($Value, '\d+')
    if ($match.Success) { return [int]$match.Value }
    return $null
}

function Get-CohortScore($File, [string]$Cohort, [string]$Job, [string]$Policy, [string]$CompanyCompact) {
    $selected = Get-CohortNumber $Cohort
    $found = @(Get-CohortNumbers $File.FullName)
    $hasSelected = $found -contains $selected
    if ($Policy -eq 'strict') {
        if ($selected -gt 1 -and -not $hasSelected) { return $null }
        if ($found.Count -gt 0 -and -not $hasSelected) { return $null }
    } elseif ($Policy -eq 'normal' -and $found.Count -gt 0 -and -not $hasSelected) {
        return $null
    }
    if ($hasSelected) { $score = 300 }
    elseif ($found.Count -eq 0) { $score = 100 }
    elseif ($Policy -eq 'common') { $score = 20 }
    else { return $null }
    $fullCompact = Compact-Text $File.FullName
    if ($fullCompact.Contains($CompanyCompact)) { $score += 1000 }
    if (-not [string]::IsNullOrWhiteSpace($Job) -and $fullCompact.Contains((Compact-Text $Job))) { $score += 100 }
    return $score
}
function New-Requirement([string]$Key, [string]$Label, [int]$Order, [scriptblock]$Matcher) {
    [PSCustomObject]@{ Key=$Key; Label=$Label; Order=$Order; Matcher=$Matcher }
}

$navy=[Drawing.Color]::FromArgb(15,23,42)
$blue=[Drawing.Color]::FromArgb(37,99,235)
$light=[Drawing.Color]::FromArgb(245,247,251)
$border=[Drawing.Color]::FromArgb(218,223,232)
$muted=[Drawing.Color]::FromArgb(100,116,139)

$form=[Windows.Forms.Form]::new()
$form.Text='기업서류 PDF 자동 수집기'
$form.ClientSize=[Drawing.Size]::new(900,840)
$form.MinimumSize=[Drawing.Size]::new(820,780)
$form.StartPosition='CenterScreen'
$form.Font=[Drawing.Font]::new('Pretendard',10)
$form.BackColor=$light
$iconPath=Join-Path $PSScriptRoot 'app_icon.ico'
if(Test-Path -LiteralPath $iconPath){$form.Icon=[Drawing.Icon]::new($iconPath)}

$root=[Windows.Forms.TableLayoutPanel]::new()
$root.Dock='Fill'; $root.ColumnCount=1; $root.RowCount=4
[void]$root.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute',92))
[void]$root.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute',466))
[void]$root.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute',68))
[void]$root.RowStyles.Add([Windows.Forms.RowStyle]::new('Percent',100))
$form.Controls.Add($root)

$header=[Windows.Forms.Panel]::new()
$header.Dock='Fill'; $header.BackColor=$navy; $header.Padding=[Windows.Forms.Padding]::new(28,18,20,12)
$title=[Windows.Forms.Label]::new()
$title.Text='기업서류 PDF 자동 수집기'; $title.ForeColor=[Drawing.Color]::White
$title.Font=[Drawing.Font]::new('Pretendard',18,[Drawing.FontStyle]::Bold); $title.Dock='Top'; $title.Height=38
$subtitle=[Windows.Forms.Label]::new()
$subtitle.Text='원본은 그대로 두고, 필요한 PDF 복사본만 새 폴더에 정리합니다.'
$subtitle.ForeColor=[Drawing.Color]::FromArgb(190,200,215); $subtitle.Dock='Top'; $subtitle.Height=24
$header.Controls.Add($subtitle); $header.Controls.Add($title); $root.Controls.Add($header,0,0)

$cardHost=[Windows.Forms.Panel]::new()
$cardHost.Dock='Fill'; $cardHost.Padding=[Windows.Forms.Padding]::new(24,20,24,8); $root.Controls.Add($cardHost,0,1)
$card=[Windows.Forms.TableLayoutPanel]::new()
$card.Dock='Fill'; $card.BackColor=[Drawing.Color]::White; $card.Padding=[Windows.Forms.Padding]::new(22,18,22,16)
$card.ColumnCount=3; $card.RowCount=9
[void]$card.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',145))
[void]$card.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent',100))
[void]$card.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',105))
1..8 | ForEach-Object { [void]$card.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute',46)) }
[void]$card.RowStyles.Add([Windows.Forms.RowStyle]::new('Percent',100))
$cardHost.Controls.Add($card)

function New-FieldLabel([string]$Text) {
    $label=[Windows.Forms.Label]::new(); $label.Text=$Text; $label.Dock='Fill'; $label.TextAlign='MiddleLeft'
    $label.Font=[Drawing.Font]::new('Pretendard',10,[Drawing.FontStyle]::Bold); return $label
}
function New-FieldBox([string]$Value='') {
    $box=[Windows.Forms.TextBox]::new(); $box.Text=$Value; $box.Dock='Fill'
    $box.Margin=[Windows.Forms.Padding]::new(0,8,10,8); $box.BorderStyle='FixedSingle'; return $box
}
function New-BrowseButton($Box) {
    $button=[Windows.Forms.Button]::new(); $button.Text='찾아보기'; $button.Dock='Fill'
    $button.Margin=[Windows.Forms.Padding]::new(0,6,0,6); $button.FlatStyle='Flat'
    $button.BackColor=[Drawing.Color]::White; $button.FlatAppearance.BorderColor=$border
    $button.Add_Click({ $p=Pick-Folder $Box.Text; if($p){$Box.Text=$p} }.GetNewClosure()); return $button
}

$defaultSearch=[Environment]::GetFolderPath('UserProfile')
$defaultDesktop=[Environment]::GetFolderPath('Desktop')
if(-not $defaultDesktop){$defaultDesktop=Join-Path $defaultSearch 'Desktop'}
$defaultOutput=Join-Path $defaultDesktop '기업별_PDF_수집결과'
$txtSearch=New-FieldBox $defaultSearch
$txtOutput=New-FieldBox $defaultOutput
$txtCompany=New-FieldBox
$txtJob=New-FieldBox
$txtParticipants=New-FieldBox
$txtMentors=New-FieldBox
$card.Controls.Add((New-FieldLabel '검색할 폴더'),0,0); $card.Controls.Add($txtSearch,1,0); $card.Controls.Add((New-BrowseButton $txtSearch),2,0)
$card.Controls.Add((New-FieldLabel '결과 저장 폴더'),0,1); $card.Controls.Add($txtOutput,1,1); $card.Controls.Add((New-BrowseButton $txtOutput),2,1)
$card.Controls.Add((New-FieldLabel '기업명'),0,2); $card.Controls.Add($txtCompany,1,2); $card.SetColumnSpan($txtCompany,2)

$cohortJob=[Windows.Forms.TableLayoutPanel]::new()
$cohortJob.Dock='Fill'; $cohortJob.ColumnCount=4
[void]$cohortJob.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',140))
[void]$cohortJob.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',70))
[void]$cohortJob.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent',100))
[void]$cohortJob.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',190))
$cmbCohort=[Windows.Forms.ComboBox]::new()
$cmbCohort.Dock='Fill'; $cmbCohort.Margin=[Windows.Forms.Padding]::new(0,7,12,7); $cmbCohort.DropDownStyle='DropDown'
[void]$cmbCohort.Items.AddRange(@('1기','2기','3기','4기','5기')); $cmbCohort.Text='1기'
$txtJob.Margin=[Windows.Forms.Padding]::new(0,8,12,8)
$jobHelp=[Windows.Forms.Label]::new(); $jobHelp.Text='선택 입력  예: 마케팅'; $jobHelp.ForeColor=$muted; $jobHelp.Dock='Fill'; $jobHelp.TextAlign='MiddleLeft'
$cohortJob.Controls.Add($cmbCohort,0,0); $cohortJob.Controls.Add((New-FieldLabel '직무'),1,0); $cohortJob.Controls.Add($txtJob,2,0); $cohortJob.Controls.Add($jobHelp,3,0)
$card.Controls.Add((New-FieldLabel '기수'),0,3); $card.Controls.Add($cohortJob,1,3); $card.SetColumnSpan($cohortJob,2)
$roundMonth=[Windows.Forms.TableLayoutPanel]::new()
$roundMonth.Dock='Fill'; $roundMonth.ColumnCount=4
[void]$roundMonth.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',140))
[void]$roundMonth.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',90))
[void]$roundMonth.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',100))
[void]$roundMonth.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent',100))
$cmbRound=[Windows.Forms.ComboBox]::new()
$cmbRound.Dock='Fill'; $cmbRound.Margin=[Windows.Forms.Padding]::new(0,7,12,7); $cmbRound.DropDownStyle='DropDownList'
[void]$cmbRound.Items.AddRange(@('1차','2차')); $cmbRound.SelectedIndex=0
$txtMonth=New-FieldBox; $txtMonth.Margin=[Windows.Forms.Padding]::new(0,8,12,8)
$monthHelp=[Windows.Forms.Label]::new(); $monthHelp.Text='예: 7 또는 6-7'; $monthHelp.ForeColor=$muted; $monthHelp.Dock='Fill'; $monthHelp.TextAlign='MiddleLeft'
$roundMonth.Controls.Add($cmbRound,0,0); $roundMonth.Controls.Add((New-FieldLabel '월 구간'),1,0)
$roundMonth.Controls.Add($txtMonth,2,0); $roundMonth.Controls.Add($monthHelp,3,0)
$card.Controls.Add((New-FieldLabel '수집 차수'),0,4); $card.Controls.Add($roundMonth,1,4); $card.SetColumnSpan($roundMonth,2)
$card.Controls.Add((New-FieldLabel '참여자명'),0,5); $card.Controls.Add($txtParticipants,1,5); $card.SetColumnSpan($txtParticipants,2)
$card.Controls.Add((New-FieldLabel '멘토명'),0,6); $card.Controls.Add($txtMentors,1,6); $card.SetColumnSpan($txtMentors,2)

$hint=[Windows.Forms.Label]::new()
$hint.Text='참여자와 멘토를 같은 순서로 쉼표 구분하세요.  예: 홍길동, 김철수'; $hint.ForeColor=$muted; $hint.Dock='Fill'; $hint.TextAlign='MiddleLeft'
$card.Controls.Add($hint,1,7); $card.SetColumnSpan($hint,2)
$monthRule=[Windows.Forms.Label]::new()
$monthRule.Text='월 구간이 일치하는 인턴형 일경험 결과보고만 수집합니다.'; $monthRule.ForeColor=$muted; $monthRule.Dock='Fill'; $monthRule.TextAlign='MiddleLeft'
$card.Controls.Add($monthRule,1,8); $card.SetColumnSpan($monthRule,2)

$actions=[Windows.Forms.FlowLayoutPanel]::new()
$actions.Dock='Fill'; $actions.Padding=[Windows.Forms.Padding]::new(170,10,0,8); $actions.WrapContents=$false
$btnRun=[Windows.Forms.Button]::new()
$btnRun.Text='PDF 복사본 수집'; $btnRun.Size=[Drawing.Size]::new(220,44)
$btnRun.BackColor=$blue; $btnRun.ForeColor=[Drawing.Color]::White; $btnRun.FlatStyle='Flat'; $btnRun.FlatAppearance.BorderSize=0
$btnOpen=[Windows.Forms.Button]::new()
$btnOpen.Text='결과 폴더 열기'; $btnOpen.Size=[Drawing.Size]::new(190,44); $btnOpen.Margin=[Windows.Forms.Padding]::new(12,0,0,0)
$btnOpen.Enabled=$false; $btnOpen.FlatStyle='Flat'; $btnOpen.BackColor=[Drawing.Color]::White; $btnOpen.FlatAppearance.BorderColor=$border
$actions.Controls.Add($btnRun); $actions.Controls.Add($btnOpen); $root.Controls.Add($actions,0,2)

$logHost=[Windows.Forms.GroupBox]::new()
$logHost.Text='수집 결과'; $logHost.Dock='Fill'; $logHost.Padding=[Windows.Forms.Padding]::new(18,18,18,16)
$logHost.Margin=[Windows.Forms.Padding]::new(24,4,24,22)
$txtLog=[Windows.Forms.TextBox]::new()
$txtLog.Dock='Fill'; $txtLog.Multiline=$true; $txtLog.ScrollBars='Vertical'; $txtLog.ReadOnly=$true
$txtLog.BackColor=[Drawing.Color]::White; $txtLog.BorderStyle='None'; $txtLog.Font=[Drawing.Font]::new('Pretendard',10)
$logHost.Controls.Add($txtLog); $root.Controls.Add($logHost,0,3)
$script:lastResultFolder = $null
$btnOpen.Add_Click({ if ($script:lastResultFolder -and (Test-Path -LiteralPath $script:lastResultFolder)) { Start-Process explorer.exe -ArgumentList ('"' + $script:lastResultFolder + '"') } })

$btnRun.Add_Click({
$searchRoot=$txtSearch.Text.Trim();$outputRoot=$txtOutput.Text.Trim();$company=(Normalize-Text $txtCompany.Text).Trim()-replace'^\(주\)',''
$cohort=$cmbCohort.Text.Trim();$job=(Normalize-Text $txtJob.Text).Trim();$round=$cmbRound.SelectedItem.ToString();$monthRange=$txtMonth.Text.Trim();$participants=@(Split-Names $txtParticipants.Text);$mentors=@(Split-Names $txtMentors.Text)
if(-not(Test-Path -LiteralPath $searchRoot -PathType Container)){[Windows.Forms.MessageBox]::Show('검색할 폴더를 확인해 주세요.');return}
if([string]::IsNullOrWhiteSpace($outputRoot)-or[string]::IsNullOrWhiteSpace($company)){[Windows.Forms.MessageBox]::Show('결과 폴더와 기업명을 입력해 주세요.');return}
if(([regex]::Matches($monthRange,'\d+')).Count-lt 1){[Windows.Forms.MessageBox]::Show('월을 7 또는 6-7과 같은 형식으로 입력해 주세요.');return}
if($cohort-notmatch'^\d+\s*기$'){[Windows.Forms.MessageBox]::Show('기수를 1기, 2기와 같은 형식으로 입력해 주세요.');return}
if($participants.Count-eq 0-or$mentors.Count-eq 0){[Windows.Forms.MessageBox]::Show('참여자명과 멘토명을 입력해 주세요.');return}

if($mentors.Count-ne 1-and$mentors.Count-ne$participants.Count){[Windows.Forms.MessageBox]::Show('멘토는 한 명만 입력하거나 참여자 수와 같게 입력해 주세요.');return}
$btnRun.Enabled=$false;$btnOpen.Enabled=$false;$txtLog.Text='PDF 파일을 검색하는 중입니다...';$form.Refresh()
try{
$allPdf=@(Get-ChildItem -LiteralPath $searchRoot -File -Recurse -Filter '*.pdf' -ErrorAction SilentlyContinue|Where-Object{-not$_.FullName.StartsWith($outputRoot,[StringComparison]::OrdinalIgnoreCase)})
$companyCompact=Compact-Text $company;$companyFiles=@($allPdf|Where-Object{(Compact-Text $_.FullName).Contains($companyCompact)})
$companyFolderName=Safe-Name ("(주)$company");$operationName=Safe-Name ((@($job,$cohort)|Where-Object{$_})-join'_');$resultFolder=Join-Path (Join-Path $outputRoot $companyFolderName) $operationName
New-Item -ItemType Directory -Path $resultFolder -Force|Out-Null
$overall=New-Object Collections.Generic.List[string];$overall.Add("기업: (주)$company");$overall.Add("기수: $cohort");$overall.Add("직무: $(if($job){$job}else{'미입력'})");$overall.Add("차수: $round");$overall.Add("월 구간: $monthRange");$overall.Add('');$totalFound=0;$totalMissing=0
for($i=0;$i-lt$participants.Count;$i++){
$p=$participants[$i];$m=if($mentors.Count-eq 1){$mentors[0]}else{$mentors[$i]}
$pc=Compact-Text $p;$mc=Compact-Text $m
$requirements=New-Object System.Collections.ArrayList
[void]$requirements.Add((New-Requirement 'implementation' '인턴형 일경험 실시보고' 1 {param($f)$c=Compact-Text $f.BaseName;(($c.StartsWith($companyCompact))-or($c.StartsWith('주'+$companyCompact)))-and($c-match'인턴형?일경험실시보고')-and-not(Is-MonthMatch $f.BaseName $monthRange)}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'monthly' "($monthRange)월 인턴형 일경험 결과보고" 2 {param($f)$c=Compact-Text $f.BaseName;(Is-MonthMatch $f.BaseName $monthRange)-and($c-match'인턴형?일경험결과보고')}))
[void]$requirements.Add((New-Requirement 'attendance' ("출석부_{0}_{1}"-f$p,$round) 3 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('출석부')-and$c.Contains($pc)-and$c.Contains($round)}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'company_bank' '기업통장사본' 4 {param($f)$c=Compact-Text $f.BaseName;$c-match'기업통장사본|국민통장사본'}))
[void]$requirements.Add((New-Requirement 'interview' "참여자($p) 멘토($m) 면담일지" 5 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('면담일지')-and$c.Contains($pc)-and$c.Contains($mc)-and($round-ne'2차'-or$c.Contains('2차'))}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'allowance' "멘토수당 신청서_$m" 6 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('멘토수당신청서')-and$c.Contains($mc)-and($round-ne'2차'-or$c.Contains('2차'))}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'mentor_bundle' "멘토($m) 증빙묶음" 7 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('멘토')-and$c.Contains($mc)-and$c.Contains('통장사본')-and$c.Contains('신분증사본')-and$c.Contains('재직증명서')}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'contract' "인턴형 프로그램 표준계약서_$p" 8 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('인턴형프로그램표준계약서')-and$c.Contains($pc)}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'participant_bundle' "참여자($p) 증빙묶음" 9 {param($f)$c=Compact-Text $f.BaseName;$evidenceCount=@('통장사본','신분증사본','주민등록등본','길찾기결과'|Where-Object{$c.Contains($_)}).Count;$c.Contains('참여자')-and$c.Contains($pc)-and($c.Contains('증빙묶음')-or$evidenceCount-ge 2)}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'privacy' "개인정보 활용에 관한 동의서_$p" 10 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('개인정보')-and$c.Contains('동의서')-and$c.Contains($pc)}.GetNewClosure()))
if($round-eq'2차'){
[void]$requirements.Add((New-Requirement 'completion_report' '인턴형 일경험 실시 결과보고' 11 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('인턴형일경험실시결과보고')}))
[void]$requirements.Add((New-Requirement 'participant_summary' "인턴형 일경험 종합 보고서_참여자($p)" 12 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('인턴형일경험종합보고서')-and$c.Contains('참여자')-and$c.Contains($pc)}.GetNewClosure()))
[void]$requirements.Add((New-Requirement 'mentor_result' "멘토($m) 결과보고서" 13 {param($f)$c=Compact-Text $f.BaseName;$c.Contains('멘토')-and$c.Contains($mc)-and$c.Contains('결과보고서')}.GetNewClosure()))
}
$roleFolder=Join-Path $resultFolder(Safe-Name $p);New-Item -ItemType Directory -Path $roleFolder -Force|Out-Null
$report=New-Object Collections.Generic.List[string];$missingLabels=New-Object Collections.Generic.List[string];$report.Add("참여자: $p");$report.Add("멘토: $m");$report.Add('');$foundCount=0;$missingCount=0
foreach($req in($requirements|Sort-Object Order)){
$companyOrders=@(1,2,4,11);$commonOrders=@(4,7);$strictOrders=@(1,2,6,11);$pool=if($companyOrders-contains$req.Order){$companyFiles}else{$allPdf};$policy=if($commonOrders-contains$req.Order){'common'}elseif($strictOrders-contains$req.Order){'strict'}else{'normal'};$ranked=@($pool|Where-Object{&$req.Matcher $_}|ForEach-Object{$score=Get-CohortScore $_ $cohort $job $policy $companyCompact;if($null-ne$score){[PSCustomObject]@{Score=$score;Modified=$_.LastWriteTime;File=$_}}}|Sort-Object Score,Modified -Descending);$matches=@($ranked|ForEach-Object{$_.File});Get-ChildItem -LiteralPath $roleFolder -File -Filter ("$($req.Order). *.pdf") -ErrorAction SilentlyContinue|Remove-Item -Force
if($matches.Count-gt 0){$chosen=$matches[0];$targetName=Safe-Name("$($req.Order). "+$chosen.Name);Copy-Item -LiteralPath $chosen.FullName -Destination(Join-Path $roleFolder $targetName)-Force;$foundCount++;$totalFound++;$report.Add("[O] $($req.Order). $($req.Label)");$report.Add("    원본: $($chosen.FullName)")}
else{$missingCount++;$totalMissing++;$missingLabels.Add("$($req.Order). $($req.Label)");$report.Add("[X] $($req.Order). $($req.Label)")}
}
$report.Add('');$report.Add("수집: $($foundCount)개 / 누락: $($missingCount)개");$overall.Add("[$p / $m] 수집 $($foundCount)개, 누락 $($missingCount)개");if($missingLabels.Count-eq 0){$overall.Add("  누락 없음")}else{foreach($missingLabel in $missingLabels){$overall.Add("  누락: $missingLabel")}};$overall.Add('')
}
$overall.Add('');$overall.Add("전체 수집: $($totalFound)개 / 전체 누락: $($totalMissing)개")
$script:lastResultFolder=$resultFolder;$btnOpen.Enabled=$true;$txtLog.Text=($overall-join[Environment]::NewLine)
[Windows.Forms.MessageBox]::Show(("참여자별 수집 완료: $($totalFound)개 / 누락 $($totalMissing)개"),'완료')|Out-Null
}catch{$txtLog.Text="오류: $($_.Exception.Message)";[Windows.Forms.MessageBox]::Show($_.Exception.Message,'오류')|Out-Null}finally{$btnRun.Enabled=$true}
})
[void]$form.ShowDialog()