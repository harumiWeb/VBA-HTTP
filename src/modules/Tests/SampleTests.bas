Attribute VB_Name = "SampleTests"
Option Explicit

' xlflow tests are public parameterless Sub procedures whose names match
' Test* or *_Test.  Parameterized tests use ByVal scalar arguments plus
' @TestCase(...) comments.
'
' Use XlflowAssert helpers to raise clear, JSON-friendly failures.
'
' Useful commands:
'   xlflow test
'   xlflow test --json
'   xlflow test --fail-fast
'   xlflow test --max-failures 3
'   xlflow test --rerun-failed 1
'
' Optional hooks named BeforeAll / AfterAll / BeforeEach / AfterEach run
' around tests in this module.  Keep tests independent; use hooks only
' when setup or cleanup is actually needed.

'@Tag("smoke")
Public Sub Test_Sample_Pass()
    XlflowAssert.AssertEquals 2, 1 + 1, "basic arithmetic should work"
    XlflowAssert.AssertTrue Len("xlflow") > 0, "strings should have length"
End Sub

'@TestCase("adds positives"; 1, 2, 3)
'@TestCase("adds negatives"; -1, -2, -3)
Public Sub Test_Adds_Numbers(ByVal leftValue As Long, ByVal rightValue As Long, ByVal expected As Long)
    XlflowAssert.AssertEquals expected, leftValue + rightValue, "sum should match"
End Sub

'@ExpectedError(5)
Public Sub Test_Expected_Error()
    Err.Raise 5, "SampleTests", "Invalid procedure call or argument"
End Sub

'@Todo("not implemented yet")
Public Sub Test_Sample_Todo()
    ' Keep planned tests visible without executing them.
End Sub
