Attribute VB_Name = "HttpBodyTests"
Option Explicit

Public Sub Test_Body_DefaultIsEmpty()
    Dim Body As New HttpBody

    XlflowAssert.AssertTrue Body.IsEmpty
    XlflowAssert.AssertEquals HttpBodyEmpty, Body.Kind
End Sub

Public Sub Test_Body_TextProducesUtf8Bytes()
    Dim Body As New HttpBody
    Dim bytes As Variant

    Body.Text = ChrW$(&HE9)
    bytes = Body.Bytes

    XlflowAssert.AssertEquals HttpBodyText, Body.Kind
    XlflowAssert.AssertEquals &HC3, HttpEncoding.ByteAt(bytes, 0)
    XlflowAssert.AssertEquals &HA9, HttpEncoding.ByteAt(bytes, 1)
End Sub

Public Sub Test_Body_BytesUsesDefensiveCopies()
    Dim original(0 To 1) As Byte
    Dim firstRead As Variant
    Dim secondRead As Variant
    Dim Body As New HttpBody

    original(0) = 10
    original(1) = 20
    Body.SetBytes original
    original(0) = 99
    firstRead = Body.Bytes
    HttpEncoding.SetByteAt firstRead, 1, 88
    secondRead = Body.Bytes

    XlflowAssert.AssertEquals 10, HttpEncoding.ByteAt(secondRead, 0)
    XlflowAssert.AssertEquals 20, HttpEncoding.ByteAt(secondRead, 1)
End Sub

Public Sub Test_Body_CloneIsIndependent()
    Dim Body As New HttpBody
    Dim copy As HttpBody

    Body.Text = "before"
    Set copy = Body.Clone()
    copy.Text = "after"

    XlflowAssert.AssertEquals "before", Body.Text
    XlflowAssert.AssertEquals "after", copy.Text
End Sub
