# Test Runner Script for Change Report Email Notifications
# Runs all unit tests and integration tests and displays results

param(
    [switch]$Verbose,
    [string]$TestName = "*",
    [switch]$UnitOnly,
    [switch]$IntegrationOnly,
    [string]$TestType = "All"  # All, Unit, Integration
)

Write-Host "Change Report Email Notifications - Test Runner" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

try {
    # Check if Pester is available
    $pesterModule = Get-Module -ListAvailable Pester
    if (-not $pesterModule) {
        Write-Error "Pester module is not installed. Please install Pester to run tests."
        Write-Host "To install Pester, run: Install-Module -Name Pester -Force -SkipPublisherCheck" -ForegroundColor Yellow
        exit 1
    }
    
    # Determine which tests to run
    $testFiles = @()
    
    if ($UnitOnly -or $TestType -eq "Unit") {
        Write-Host "Running Unit Tests Only..." -ForegroundColor Green
        $testFiles = @(
            "tests\Config.Tests.ps1",
            "tests\Database.Tests.ps1", 
            "tests\Email.Tests.ps1",
            "tests\Logging.Tests.ps1"
        )
    }
    elseif ($IntegrationOnly -or $TestType -eq "Integration") {
        Write-Host "Running Integration Tests Only..." -ForegroundColor Green
        $testFiles = @("tests\Integration.Tests.ps1")
    }
    else {
        Write-Host "Running All Tests (Unit + Integration + Task Scheduler)..." -ForegroundColor Green
        $testFiles = @(
            "tests\Config.Tests.ps1",
            "tests\Database.Tests.ps1", 
            "tests\Email.Tests.ps1",
            "tests\Logging.Tests.ps1",
            "tests\Integration.Tests.ps1",
            "tests\TaskScheduler.Tests.ps1"
        )
    }
    
    # Verify test files exist
    $missingFiles = @()
    foreach ($testFile in $testFiles) {
        if (-not (Test-Path $testFile)) {
            $missingFiles += $testFile
        }
    }
    
    if ($missingFiles.Count -gt 0) {
        Write-Warning "The following test files were not found:"
        $missingFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        Write-Host "Continuing with available test files..." -ForegroundColor Yellow
        $testFiles = $testFiles | Where-Object { Test-Path $_ }
    }
    
    if ($testFiles.Count -eq 0) {
        Write-Error "No test files found to execute."
        exit 1
    }
    
    # Set test parameters
    $testParams = @{
        Path = $testFiles
        PassThru = $true
    }
    
    if ($Verbose) {
        $testParams.Verbose = $true
    }
    
    if ($TestName -ne "*") {
        $testParams.TestName = $TestName
    }
    
    # Display test execution plan
    Write-Host "`nTest Execution Plan:" -ForegroundColor Yellow
    Write-Host "===================" -ForegroundColor Yellow
    $testFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
    Write-Host ""
    
    # Run the tests
    Write-Host "Executing Tests..." -ForegroundColor Cyan
    $results = Invoke-Pester @testParams
    
    # Display detailed summary
    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    Write-Host "TEST EXECUTION SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan
    
    Write-Host "Total Tests: $($results.TotalCount)" -ForegroundColor White
    Write-Host "Passed: $($results.PassedCount)" -ForegroundColor Green
    Write-Host "Failed: $($results.FailedCount)" -ForegroundColor $(if ($results.FailedCount -gt 0) { "Red" } else { "Green" })
    Write-Host "Skipped: $($results.SkippedCount)" -ForegroundColor Yellow
    Write-Host "Execution Time: $($results.Time)" -ForegroundColor White
    
    # Calculate success rate
    if ($results.TotalCount -gt 0) {
        $successRate = [math]::Round(($results.PassedCount / $results.TotalCount) * 100, 2)
        Write-Host "Success Rate: $successRate%" -ForegroundColor $(if ($successRate -eq 100) { "Green" } else { "Yellow" })
    }
    
    # Display test results by category
    if ($results.TestResult) {
        $testsByFile = $results.TestResult | Group-Object { 
            $testPath = $_.Describe
            if ($testPath -match "Integration Tests") { "Integration Tests" }
            elseif ($testPath -match "Config") { "Configuration Tests" }
            elseif ($testPath -match "Database") { "Database Tests" }
            elseif ($testPath -match "Email") { "Email Tests" }
            elseif ($testPath -match "Logging") { "Logging Tests" }
            else { "Other Tests" }
        }
        
        Write-Host "`nResults by Test Category:" -ForegroundColor Yellow
        Write-Host "-" * 30 -ForegroundColor Yellow
        
        foreach ($category in $testsByFile) {
            $categoryPassed = ($category.Group | Where-Object { $_.Result -eq "Passed" }).Count
            $categoryFailed = ($category.Group | Where-Object { $_.Result -eq "Failed" }).Count
            $categorySkipped = ($category.Group | Where-Object { $_.Result -eq "Skipped" }).Count
            $categoryTotal = $category.Count
            
            $categoryColor = if ($categoryFailed -eq 0) { "Green" } else { "Red" }
            Write-Host "$($category.Name): $categoryPassed/$categoryTotal passed" -ForegroundColor $categoryColor
            
            if ($categoryFailed -gt 0) {
                Write-Host "  Failed: $categoryFailed" -ForegroundColor Red
            }
            if ($categorySkipped -gt 0) {
                Write-Host "  Skipped: $categorySkipped" -ForegroundColor Yellow
            }
        }
    }
    
    # Display failed tests with details
    if ($results.FailedCount -gt 0) {
        Write-Host "`nFAILED TESTS DETAILS:" -ForegroundColor Red
        Write-Host "-" * 30 -ForegroundColor Red
        
        $failedTests = $results.TestResult | Where-Object { $_.Result -eq "Failed" }
        foreach ($failedTest in $failedTests) {
            Write-Host "❌ $($failedTest.Describe) - $($failedTest.Name)" -ForegroundColor Red
            Write-Host "   Error: $($failedTest.FailureMessage)" -ForegroundColor DarkRed
            if ($failedTest.ErrorRecord) {
                Write-Host "   Location: $($failedTest.ErrorRecord.InvocationInfo.ScriptName):$($failedTest.ErrorRecord.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkRed
            }
            Write-Host ""
        }
        
        Write-Host "RECOMMENDATION: Review failed tests and fix issues before deployment." -ForegroundColor Yellow
        exit 1
    } 
    else {
        Write-Host "`n✅ ALL TESTS PASSED SUCCESSFULLY!" -ForegroundColor Green
        Write-Host "The Change Report Email Notification system is ready for deployment." -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Error "Error running tests: $($_.Exception.Message)"
    exit 1
}