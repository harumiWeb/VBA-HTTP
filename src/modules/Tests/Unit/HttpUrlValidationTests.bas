Attribute VB_Name = "HttpUrlValidationTests"
Option Explicit

Public Sub Test_UrlValidation_RemovesFragmentFromRequestTarget()
    XlflowAssert.AssertEquals "https://example.test/items?q=1", _
        HttpUrlValidation.WithoutFragment("https://example.test/items?q=1#section")
End Sub

Public Sub Test_UrlValidation_PreservesEncodedFragmentMarker()
    XlflowAssert.AssertEquals "https://example.test/items?q=%23section", _
        HttpUrlValidation.WithoutFragment("https://example.test/items?q=%23section")
End Sub
