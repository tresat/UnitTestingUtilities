Option Explicit On
Option Strict On

Imports NUnit.Framework
Imports System.Reflection

''' <summary>
''' This class exists to serve as a base class for all unit test classes.  It provides useful
''' utilities (like private method callers, private field accessors) which are useful in
''' testing.
'''
''' Created by: tom on 5/1/2009 3:21 PM
''' </summary>
Public Class clsBaseUnitTester
    ''' <summary>
    ''' Sets the value of a private field.
    ''' </summary>
    ''' <param name="pobjInstance">An instance of the class.</param>
    ''' <param name="pstrFieldName">Name of the field.</param>
    ''' <param name="pobjValue">The value to set the field to.</param>
    Protected Sub SetPrivateFieldValue(ByVal pobjInstance As Object, ByVal pstrFieldName As String, ByRef pobjValue As Object)

        Dim t As Type = pobjInstance.GetType()
        Dim fi As FieldInfo = t.GetField(pstrFieldName, BindingFlags.Instance Or BindingFlags.NonPublic)

        fi.SetValue(pobjInstance, pobjValue, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, System.Globalization.CultureInfo.CurrentCulture)
    End Sub


    ''' <summary>
    ''' Gets the value of a private field in a class.
    ''' </summary>
    ''' <param name="pobjInstance">An instance of the class.</param>
    ''' <param name="pstrFieldName">Name of the private field.</param>
    ''' <returns>The value of the field.</returns>
    Protected Function GetPrivateFieldValue(ByVal pobjInstance As Object, ByVal pstrFieldName As String) As Object

        Dim t As Type = pobjInstance.GetType()
        Dim fi As FieldInfo = t.GetField(pstrFieldName, BindingFlags.Instance Or BindingFlags.NonPublic)

        Return fi.GetValue(pobjInstance)
    End Function

    ''' <summary>
    ''' Calls a private method of a class.
    ''' </summary>
    ''' <param name="pobjInstance">An instance of the class.</param>
    ''' <param name="pstrMethodName">Name of the private method.</param>
    ''' <param name="paobjParamList">The param list for the method, as an object array.</param>
    ''' <returns>The result of the function call.</returns>
    Protected Function CallPrivateMethod(ByVal pobjInstance As Object, ByVal pstrMethodName As String, Optional ByVal paobjParamList As Object() = Nothing) As Object

        Dim t As Type = pobjInstance.GetType()
        Dim mi As MethodInfo

        If paobjParamList IsNot Nothing Then
            Dim paramTypes(paobjParamList.Count) As Type

            For intIdx As Integer = 0 To paobjParamList.Count - 1
                paramTypes(intIdx) = paobjParamList(intIdx).GetType()
            Next

            mi = t.GetMethod(pstrMethodName, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, paramTypes, Nothing)

            Return mi.Invoke(pobjInstance, paobjParamList)
        Else
            mi = t.GetMethod(pstrMethodName, BindingFlags.Instance Or BindingFlags.NonPublic)


            Return mi.Invoke(pobjInstance, Nothing)
        End If
    End Function
End Class