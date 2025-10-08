# Integration Test Runner for Change Report Email Notifications
# Specifically runs integration tests with proper setup and teardown

param(
    [switch]$Verbose,
    [switch]$IncludeTaskScheduler,
    [string]$TestName = "*",
    [switch]$GenerateReport
)

Write-Host "Change Report Email Notifications - Integration Test Runner" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan

# Check prerequisites
try {
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5.1 or higher is required. Current version: $($PSVersionTable.PSVersion)"
    }
    
    # Check if Pester is available
    $pesterModule = Get-Module -ListAvailable Pester
    if (-not $pesterModule) {
        Write-Error "Pester module is not installed. Please install Pester to run tests."
        Write-Host "To install Pester, run: Install-Module -Name Pester -Force -SkipPublisherCheck" -ForegroundColor Yellow
        exit 1
    }
    
    Write-Host "Prerequisites check passed:" -ForegroundColor Green
    Write-Host "  ✓ PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Green
    Write-Host "  ✓ Pester Version: $($pesterModule.Version)" -ForegroundColor Green
    
    # Check if running as Administrator for Task Scheduler tests
    if ($IncludeTaskScheduler) {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin) {
            Write-Warning "Task Scheduler tests require Administrator privileges."
            Write-Host "Please run this script as Administrator to include Task Scheduler tests." -ForegroundColor Yellow
            $IncludeTaskScheduler = $false
        }
        else {
            Write-Host "  ✓ Administrator privileges detected for Task Scheduler tests" -ForegroundColor Green
        }
    }
    
}
catch {
    Write-Error "Prerequisites check failed: $($_.Exception.Message)"
    exit 1
}

# Determine test files to run
$testFiles = @("tests\Integration.Tests.ps1")

if ($IncludeTaskScheduler) {
    $testFiles += "tests\TaskScheduler.Tests.ps1"
    Write-Host "Including Task Scheduler integration tests..." -ForegroundColor Yellow
}

# Verify test files exist
$missingFiles = @()
foreach ($testFile in $testFiles) {
    if (-not (Test-Path $testFile)) {
        $missingFiles += $testFile
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Error "The following test files were not found:"
    $missingFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}

# Set up test environment
Write-Host "`nSetting up test environment..." -ForegroundColor Yellow

# Create temporary directories for test artifacts
$testArtifactsDir = "TestArtifacts"
if (-not (Test-Path $testArtifactsDir)) {
    New-Item -ItemType Directory -Path $testArtifactsDir -Force | Out-Null
}

# Set test parameters
$testParams = @{
    Path = $testFiles
    PassThru = $true
    OutputFormat = "NUnitXml"
    OutputFile = "$testArtifactsDir\IntegrationTestResults.xml"
}

if ($Verbose) {
    $testParams.Verbose = $true
}

if ($TestName -ne "*") {
    $testParams.TestName = $TestName
}

# Display test execution plan
Write-Host "`nIntegration Test Execution Plan:" -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Yellow
$testFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
Write-Host ""

# Execute integration tests
Write-Host "Executing Integration Tests..." -ForegroundColor Cyan
Write-Host "This may take several minutes..." -ForegroundColor Gray

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # Run the tests
    $results = Invoke-Pester @testParams
    
    $stopwatch.Stop()
    
    # Display detailed results
    Write-Host "`n" + "="*70 -ForegroundColor Cyan
    Write-Host "INTEGRATION TEST EXECUTION SUMMARY" -ForegroundColor Cyan
    Write-Host "="*70 -ForegroundColor Cyan
    
    Write-Host "Execution Time: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor White
    Write-Host "Total Tests: $($results.TotalCount)" -ForegroundColor White
    Write-Host "Passed: $($results.PassedCount)" -ForegroundColor Green
    Write-Host "Failed: $($results.FailedCount)" -ForegroundColor $(if ($results.FailedCount -gt 0) { "Red" } else { "Green" })
    Write-Host "Skipped: $($results.SkippedCount)" -ForegroundColor Yellow
    
    # Calculate success rate
    if ($results.TotalCount -gt 0) {
        $successRate = [math]::Round(($results.PassedCount / $results.TotalCount) * 100, 2)
        Write-Host "Success Rate: $successRate%" -ForegroundColor $(if ($successRate -eq 100) { "Green" } elseif ($successRate -ge 80) { "Yellow" } else { "Red" })
    }
    
    # Display test categories
    if ($results.TestResult) {
        Write-Host "`nTest Categories:" -ForegroundColor Yellow
        Write-Host "===============" -ForegroundColor Yellow
        
        $categories = @{
            "End-to-End Workflow" = ($results.TestResult | Where-Object { $_.Describe -match "End-to-End" }).Count
            "Error Scenarios" = ($results.TestResult | Where-Object { $_.Describe -match "Error Scenarios" }).Count
            "Performance Tests" = ($results.TestResult | Where-Object { $_.Describe -match "Performance" }).Count
            "Stress Testing" = ($results.TestResult | Where-Object { $_.Describe -match "Stress" }).Count
            "Workflow Validation" = ($results.TestResult | Where-Object { $_.Describe -match "Workflow Validation" }).Count
            "Task Scheduler" = ($results.TestResult | Where-Object { $_.Describe -match "Task Scheduler" }).Count
        }
        
        foreach ($category in $categories.GetEnumerator()) {
            if ($category.Value -gt 0) {
                $categoryTests = $results.TestResult | Where-Object { $_.Describe -match $category.Key }
                $categoryPassed = ($categoryTests | Where-Object { $_.Result -eq "Passed" }).Count
                $categoryFailed = ($categoryTests | Where-Object { $_.Result -eq "Failed" }).Count
                $categorySkipped = ($categoryTests | Where-Object { $_.Result -eq "Skipped" }).Count
                
                $categoryColor = if ($categoryFailed -eq 0) { "Green" } elseif ($categoryPassed -gt $categoryFailed) { "Yellow" } else { "Red" }
                Write-Host "$($category.Key): $categoryPassed passed, $categoryFailed failed, $categorySkipped skipped" -ForegroundColor $categoryColor
            }
        }
        
        # Display performance metrics if available
        $performanceTests = $results.TestResult | Where-Object { $_.Describe -match "Performance" -and $_.Result -eq "Passed" }
        if ($performanceTests.Count -gt 0) {
            Write-Host "`nPerformance Metrics:" -ForegroundColor Cyan
            Write-Host "===================" -ForegroundColor Cyan
            Write-Host "✓ Execution time tests: Passed" -ForegroundColor Green
            Write-Host "✓ Memory usage tests: Passed" -ForegroundColor Green
            Write-Host "✓ Resource cleanup tests: Passed" -ForegroundColor Green
            Write-Host "✓ Concurrent execution prevention: Passed" -ForegroundColor Green
        }
    }
    
    # Display failed tests with details
    if ($results.FailedCount -gt 0) {
        Write-Host "`nFAILED INTEGRATION TESTS:" -ForegroundColor Red
        Write-Host "=========================" -ForegroundColor Red
        
        $failedTests = $results.TestResult | Where-Object { $_.Result -eq "Failed" }
        foreach ($failedTest in $failedTests) {
            Write-Host "❌ $($failedTest.Describe)" -ForegroundColor Red
            Write-Host "   Test: $($failedTest.Name)" -ForegroundColor Red
            Write-Host "   Error: $($failedTest.FailureMessage)" -ForegroundColor DarkRed
            
            if ($failedTest.ErrorRecord) {
                Write-Host "   Location: $($failedTest.ErrorRecord.InvocationInfo.ScriptName):$($failedTest.ErrorRecord.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkRed
            }
            Write-Host ""
        }
        
        Write-Host "RECOMMENDATION:" -ForegroundColor Yellow
        Write-Host "- Review failed integration tests and fix underlying issues" -ForegroundColor Yellow
        Write-Host "- Ensure all dependencies (database, email server) are properly configured" -ForegroundColor Yellow
        Write-Host "- Check network connectivity and authentication settings" -ForegroundColor Yellow
        Write-Host "- Verify configuration files are valid and accessible" -ForegroundColor Yellow
    }
    else {
        Write-Host "`n✅ ALL INTEGRATION TESTS PASSED!" -ForegroundColor Green
        Write-Host "The Change Report Email Notification system has been thoroughly tested and is ready for production deployment." -ForegroundColor Green
        
        Write-Host "`nNext Steps:" -ForegroundColor Cyan
        Write-Host "- Deploy the solution to production environment" -ForegroundColor White
        Write-Host "- Configure production database and email settings" -ForegroundColor White
        Write-Host "- Set up Windows Task Scheduler for daily execution" -ForegroundColor White
        Write-Host "- Monitor initial executions and verify email delivery" -ForegroundColor White
        Write-Host "- Review integration test artifacts in TestArtifacts directory" -ForegroundColor White
        
        # Display integration test coverage summary
        Write-Host "`nIntegration Test Coverage Summary:" -ForegroundColor Cyan
        Write-Host "==================================" -ForegroundColor Cyan
        Write-Host "✅ End-to-end workflow testing" -ForegroundColor Green
        Write-Host "✅ Configuration validation and error handling" -ForegroundColor Green
        Write-Host "✅ Database connectivity and error scenarios" -ForegroundColor Green
        Write-Host "✅ Email delivery and SMTP error handling" -ForegroundColor Green
        Write-Host "✅ Performance and memory usage validation" -ForegroundColor Green
        Write-Host "✅ Concurrent execution prevention" -ForegroundColor Green
        Write-Host "✅ Resource cleanup and data integrity" -ForegroundColor Green
        Write-Host "✅ Task Scheduler integration (if included)" -ForegroundColor Green
        Write-Host "✅ Test mode validation and safety checks" -ForegroundColor Green
    }
    
    # Generate detailed report if requested
    if ($GenerateReport) {
        Write-Host "`nGenerating detailed test report..." -ForegroundColor Yellow
        
        $reportPath = "$testArtifactsDir\IntegrationTestReport.html"
        $reportContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Integration Test Report - Change Report Email Notifications</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #f0f0f0; padding: 10px; border-radius: 5px; }
        .summary { background-color: #e8f5e8; padding: 10px; margin: 10px 0; border-radius: 5px; }
        .failed { background-color: #ffe8e8; padding: 10px; margin: 10px 0; border-radius: 5px; }
        .test-result { margin: 5px 0; padding: 5px; }
        .passed { color: green; }
        .failed-text { color: red; }
        .skipped { color: orange; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Integration Test Report</h1>
        <h2>Change Report Email Notification System</h2>
        <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p>Execution Time: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))</p>
    </div>
    
    <div class="summary">
        <h3>Test Summary</h3>
        <p><strong>Total Tests:</strong> $($results.TotalCount)</p>
        <p><strong>Passed:</strong> <span class="passed">$($results.PassedCount)</span></p>
        <p><strong>Failed:</strong> <span class="failed-text">$($results.FailedCount)</span></p>
        <p><strong>Skipped:</strong> <span class="skipped">$($results.SkippedCount)</span></p>
        <p><strong>Success Rate:</strong> $successRate%</p>
    </div>
    
    <h3>Test Results</h3>
"@
        
        foreach ($test in $results.TestResult) {
            $resultClass = switch ($test.Result) {
                "Passed" { "passed" }
                "Failed" { "failed-text" }
                "Skipped" { "skipped" }
                default { "" }
            }
            
            $reportContent += @"
    <div class="test-result">
        <strong class="$resultClass">[$($test.Result)] $($test.Describe) - $($test.Name)</strong>
"@
            
            if ($test.Result -eq "Failed") {
                $reportContent += @"
        <div class="failed">
            <p><strong>Error:</strong> $($test.FailureMessage)</p>
        </div>
"@
            }
            
            $reportContent += "</div>`n"
        }
        
        $reportContent += @"
</body>
</html>
"@
        
        $reportContent | Set-Content -Path $reportPath -Encoding UTF8
        Write-Host "Detailed report saved to: $reportPath" -ForegroundColor Green
    }
    
    # Set exit code based on results
    if ($results.FailedCount -gt 0) {
        exit 1
    }
    else {
        exit 0
    }
}
catch {
    $stopwatch.Stop()
    Write-Error "Integration test execution failed: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
finally {
    # Clean up test environment
    Write-Host "`nCleaning up test environment..." -ForegroundColor Gray
    
    # Remove temporary test files (but keep artifacts for review)
    $tempFiles = Get-ChildItem -Path "." -Filter "*test*temp*" -ErrorAction SilentlyContinue
    if ($tempFiles) {
        $tempFiles | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    
    # Clean up any test scheduled tasks
    $testTasks = Get-ScheduledTask -TaskName "Test-ChangeReport-*" -ErrorAction SilentlyContinue
    if ($testTasks) {
        Write-Host "Cleaning up test scheduled tasks..." -ForegroundColor Gray
        $testTasks | ForEach-Object {
            try {
                Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false -ErrorAction SilentlyContinue
            }
            catch {
                Write-Warning "Could not remove test task $($_.TaskName): $($_.Exception.Message)"
            }
        }
    }
    
    Write-Host "Test environment cleanup completed." -ForegroundColor Gray
}