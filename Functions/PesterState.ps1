function New-PesterState
{
    param (
        [String[]]$TagFilter,
        [String[]]$ExcludeTagFilter,
        [String[]]$TestNameFilter,
        [System.Management.Automation.SessionState]$SessionState,
        [Switch]$Strict,
        [Switch]$Quiet
    )

    if ($null -eq $SessionState) { $SessionState = $ExecutionContext.SessionState }

    & $SafeCommands['New-Module'] -Name Pester -AsCustomObject -ScriptBlock {
        param (
            [String[]]$_tagFilter,
            [String[]]$_excludeTagFilter,
            [String[]]$_testNameFilter,
            [System.Management.Automation.SessionState]$_sessionState,
            [Switch]$Strict,
            [Switch]$Quiet
        )

        #public read-only
        $TagFilter = $_tagFilter
        $ExcludeTagFilter = $_excludeTagFilter
        $TestNameFilter = $_testNameFilter

        $script:SessionState = $_sessionState
        $script:PesterStack = New-Object System.Collections.Stack
        $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $script:MostRecentTimestamp = 0
        $script:CommandCoverage = @()
        $script:BeforeEach = @()
        $script:AfterEach = @()
        $script:BeforeAll = @()
        $script:AfterAll = @()
        $script:Strict = $Strict
        $script:Quiet = $Quiet

        $script:TestResult = @()
        $script:TestTree = @()

        $script:TotalCount = 0
        $script:Time = [timespan]0
        $script:PassedCount = 0
        $script:FailedCount = 0
        $script:SkippedCount = 0
        $script:PendingCount = 0
        $script:InconclusiveCount = 0

        $script:SafeCommands = @{}

        $script:SafeCommands['New-Object']          = & (Pester\SafeGetCommand) -Name New-Object          -Module Microsoft.PowerShell.Utility -CommandType Cmdlet
        $script:SafeCommands['Select-Object']       = & (Pester\SafeGetCommand) -Name Select-Object       -Module Microsoft.PowerShell.Utility -CommandType Cmdlet
        $script:SafeCommands['Export-ModuleMember'] = & (Pester\SafeGetCommand) -Name Export-ModuleMember -Module Microsoft.PowerShell.Core    -CommandType Cmdlet
        $script:SafeCommands['Add-Member']          = & (Pester\SafeGetCommand) -Name Add-Member          -Module Microsoft.PowerShell.Utility -CommandType Cmdlet

        function Assert-StackFrameInProgress
        {
            param (
                [Parameter(Mandatory = $true)]
                [string] $From,

                [Parameter(Mandatory = $true)]
                [ValidateSet('TestGroup', 'TestCase')]
                [string] $FrameType
            )

            if ($script:PesterStack.Count -le 0)
            {
                throw "$From called when test stack was empty."
            }

            $frame = $script:PesterStack.Peek()

            if ($frame -isnot [System.Collections.IDictionary] -or -not $frame.Contains('Type'))
            {
                throw "$From encountered an invalid test stack frame."
            }

            if ($frame['Type'] -ne $FrameType)
            {
                throw "$From called when the current test stack frame was not a $FrameType frame.  Current frame type: '$($frame['Type'])'"
            }

        }

        function EnterTestGroup([string] $Name, [string] $TypeHint, [string[]] $Tags = @())
        {
            $frame = @{
                Name     = $Name
                TypeHint = $TypeHint
                Type     = 'TestGroup'
                Tags     = $Tags
            }

            $script:PesterStack.Push($frame)
        }

        function LeaveTestGroup()
        {
            Assert-StackFrameInProgress -From 'PesterState.LeaveTestGroup()' -FrameType TestGroup

            $null = $script:PesterStack.Pop()
        }

        function EnterTest([string]$Name)
        {
            Assert-StackFrameInProgress -From 'PesterState.EnterTest()' -FrameType TestGroup

            $frame = @{
                Name     = $Name
                TypeHint = $TypeHint
                Type     = 'TestCase'
                Tags     = $Tags
            }

            $script:PesterStack.Push($frame)
        }

        function LeaveTest
        {
            Assert-StackFrameInProgress -From 'PesterState.LeaveTest()' -FrameType TestCase

            $null = $script:PesterStack.Pop()
        }

        function AddTestResult
        {
            param (
                [string]$Name,
                [ValidateSet("Failed","Passed","Skipped","Pending","Inconclusive")]
                [string]$Result,
                [Nullable[TimeSpan]]$Time,
                [string]$FailureMessage,
                [string]$StackTrace,
                [string] $ParameterizedSuiteName,
                [System.Collections.IDictionary] $Parameters,
                [System.Management.Automation.ErrorRecord] $ErrorRecord
            )

            $previousTime = $script:MostRecentTimestamp
            $script:MostRecentTimestamp = $script:Stopwatch.Elapsed

            if ($null -eq $Time)
            {
                $Time = $script:MostRecentTimestamp - $previousTime
            }

            if (-not $script:Strict)
            {
                $Passed = "Passed","Skipped","Pending" -contains $Result
            }
            else
            {
                $Passed = $Result -eq "Passed"
                if (($Result -eq "Skipped") -or ($Result -eq "Pending"))
                {
                    $FailureMessage = "The test failed because the test was executed in Strict mode and the result '$result' was translated to Failed."
                    $Result = "Failed"
                }

            }

            $script:TotalCount++
            $script:Time += $Time

            switch ($Result)
            {
                Passed  { $script:PassedCount++; break; }
                Failed  { $script:FailedCount++; break; }
                Skipped { $script:SkippedCount++; break; }
                Pending { $script:PendingCount++; break; }
                Inconclusive { $script:InconclusiveCount++; break; }
            }

            $Script:TestResult += & $SafeCommands['New-Object'] -TypeName PsObject -Property @{
                Describe               = $CurrentDescribe
                Context                = $CurrentContext
                Name                   = $Name
                Passed                 = $Passed
                Result                 = $Result
                Time                   = $Time
                FailureMessage         = $FailureMessage
                StackTrace             = $StackTrace
                ErrorRecord            = $ErrorRecord
                ParameterizedSuiteName = $ParameterizedSuiteName
                Parameters             = $Parameters
                Quiet                  = $script:Quiet
            } | & $SafeCommands['Select-Object'] Describe, Context, Name, Result, Passed, Time, FailureMessage, StackTrace, ErrorRecord, ParameterizedSuiteName, Parameters
        }

        $ExportedVariables = "TagFilter",
        "ExcludeTagFilter",
        "TestNameFilter",
        "TestResult",
        "CurrentContext",
        "CurrentDescribe",
        "CurrentTest",
        "SessionState",
        "CommandCoverage",
        "BeforeEach",
        "AfterEach",
        "BeforeAll",
        "AfterAll",
        "Strict",
        "Quiet",
        "Time",
        "TotalCount",
        "PassedCount",
        "FailedCount",
        "SkippedCount",
        "PendingCount",
        "InconclusiveCount"

        $ExportedFunctions = "EnterContext",
        "LeaveContext",
        "EnterDescribe",
        "LeaveDescribe",
        "EnterTest",
        "LeaveTest",
        "AddTestResult"

        & $SafeCommands['Export-ModuleMember'] -Variable $ExportedVariables -function $ExportedFunctions
    } -ArgumentList $TagFilter, $ExcludeTagFilter, $TestNameFilter, $SessionState, $Strict, $Quiet |
    & $SafeCommands['Add-Member'] -MemberType ScriptProperty -Name Scope -Value {
        if ($this.CurrentTest) { 'It' }
        elseif ($this.CurrentContext)  { 'Context' }
        elseif ($this.CurrentDescribe) { 'Describe' }
        else { $null }
    } -Passthru |
    & $SafeCommands['Add-Member'] -MemberType ScriptProperty -Name ParentScope -Value {
        $parentScope = $null
        $scope = $this.Scope

        if ($scope -eq 'It' -and $this.CurrentContext)
        {
            $parentScope = 'Context'
        }

        if ($null -eq $parentScope -and $scope -ne 'Describe' -and $this.CurrentDescribe)
        {
            $parentScope = 'Describe'
        }

        return $parentScope
    } -PassThru
}
