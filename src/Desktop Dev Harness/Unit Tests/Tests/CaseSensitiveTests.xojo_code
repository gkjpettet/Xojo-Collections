#tag Class
Protected Class CaseSensitiveTests
Inherits TestGroup
	#tag Event
		Sub Setup()
		  
		End Sub
	#tag EndEvent

	#tag Event
		Function UnhandledException(err As RuntimeException, methodName As Text) As Boolean
		  #pragma unused err
		  
		  Const kMethodName As Text = "UnhandledException"
		  
		  If methodName.Length >= kMethodName.Length And methodName.Left(kMethodName.Length) = kMethodName Then
		    Assert.Pass("Exception was handled")
		    Return True
		  End If
		End Function
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Function AsText(s As String) As Text
		  Return s.ToText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CloneTest()
		  Var clone As BetterDictionary = TestDictionary.Clone
		  
		  For Each entry As DictionaryEntry In TestDictionary
		    Assert.IsTrue(clone.HasKey(entry.Key))
		    Assert.IsTrue(clone.Value(entry.key) = TestDictionary.Value(entry.Key))
		  Next entry
		  
		  Assert.IsTrue(TestDictionary.CaseSensitive = clone.CaseSensitive)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ConstructorTest()
		  Var bd As New BetterDictionary( _
		  True, _
		  AsText("lowercase text") : "lowercase text:value", _
		  AsText("MixedCase Text") : "MixedCase Text:value", _
		  AsText("UPPERCASE TEXT") : "UPPERCASE TEXT:value", _
		  AsText("lowercase string") : "lowercase string:value", _
		  AsText("MixedCase String") : "MixedCase String:value", _
		  AsText("UPPERCASE STRING") : "UPPERCASE STRING:value")
		  
		  Assert.IsTrue(bd.HasKey("lowercase text"))
		  Assert.IsTrue(bd.value("lowercase text") = "lowercase text:value")
		  
		  Assert.IsTrue(bd.HasKey("MixedCase Text"))
		  Assert.IsTrue(bd.value("MixedCase Text") = "MixedCase Text:value")
		  
		  Assert.IsTrue(bd.HasKey("UPPERCASE TEXT"))
		  Assert.IsTrue(bd.value("UPPERCASE TEXT") = "UPPERCASE TEXT:value")
		  
		  Assert.IsTrue(bd.HasKey("lowercase string"))
		  Assert.IsTrue(bd.value("lowercase string") = "lowercase string:value")
		  
		  Assert.IsTrue(bd.HasKey("MixedCase String"))
		  Assert.IsTrue(bd.value("MixedCase String") = "MixedCase String:value")
		  
		  Assert.IsTrue(bd.HasKey("UPPERCASE STRING"))
		  Assert.IsTrue(bd.value("UPPERCASE STRING") = "UPPERCASE STRING:value")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ConstructTestDictionary() As BetterDictionary
		  Var bd As New BetterDictionary(True)
		  
		  bd.Value(AsText("lowercase text")) = "lowercase text:value"
		  bd.Value(AsText("MixedCase Text")) = "MixedCase Text:value"
		  bd.Value("lowercase string") = "lowercase string:value"
		  bd.Value("MixedCase String") = "MixedCase String:value"
		  bd.Value(AsText("UPPERCASE TEXT")) = "UPPERCASE TEXT:value"
		  bd.Value("UPPERCASE STRING") = "UPPERCASE STRING:value"
		  
		  bd.Value(1) = "one"
		  bd.Value(True) = "True"
		  
		  Return bd
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasKeyTest()
		  Var textKey1 As Text = "lowercase text"
		  Var textKey2 As Text = "MixedCase Text"
		  Var textKey3 As Text = "UPPERCASE TEXT"
		  Var stringKey1 As Text = "lowercase string"
		  Var stringKey2 As Text = "MixedCase String"
		  Var stringKey3 As Text = "UPPERCASE STRING"
		  
		  Assert.IsTrue(TestDictionary.HasKey(1) = True)
		  Assert.IsTrue(TestDictionary.HasKey(True) = True)
		  
		  // Test using variables.
		  Assert.IsTrue(TestDictionary.HasKey(textKey1))
		  Assert.IsTrue(TestDictionary.HasKey(textKey2))
		  Assert.IsTrue(TestDictionary.HasKey(textKey3))
		  Assert.IsTrue(TestDictionary.HasKey(stringKey1))
		  Assert.IsTrue(TestDictionary.HasKey(stringKey2))
		  Assert.IsTrue(TestDictionary.HasKey(stringKey3))
		  
		  // Test using IDE string literals.
		  Assert.IsTrue(TestDictionary.HasKey("lowercase text"))
		  Assert.IsTrue(TestDictionary.HasKey("MixedCase Text"))
		  Assert.IsTrue(TestDictionary.HasKey("UPPERCASE TEXT"))
		  Assert.IsTrue(TestDictionary.HasKey("lowercase string"))
		  Assert.IsTrue(TestDictionary.HasKey("MixedCase String"))
		  Assert.IsTrue(TestDictionary.HasKey("UPPERCASE STRING"))
		  
		  // Fails.
		  Assert.IsFalse(TestDictionary.HasKey("LOWERCASE TEXT"))
		  Assert.IsFalse(TestDictionary.HasKey("uppercase string"))
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub IteratorTest()
		  For Each entry As DictionaryEntry In TestDictionary
		    Assert.IsTrue(TestDictionary.HasKey(entry.Key))
		    Assert.IsTrue(TestDictionary.Value(entry.Key) = entry.Value)
		  Next entry
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub KeysTest()
		  Var keys() As Variant = TestDictionary.Keys
		  For Each key As Variant In keys
		    Assert.IsTrue(TestDictionary.HasKey(key))
		  Next key
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub KeyTest()
		  Var keys() As Variant = TestDictionary.Keys
		  
		  For i As Integer = 0 To keys.LastIndex
		    Assert.IsTrue(TestDictionary.Key(i) = keys(i))
		  Next i
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LookupTest()
		  Var textKey1 As Text = "lowercase text"
		  Var textKey2 As Text = "MixedCase Text"
		  Var textKey3 As Text = "UPPERCASE TEXT"
		  Var stringKey1 As Text = "lowercase string"
		  Var stringKey2 As Text = "MixedCase String"
		  Var stringKey3 As Text = "UPPERCASE STRING"
		  
		  // Test using variables.
		  Assert.IsTrue(TestDictionary.Lookup(textKey1, 1) = "lowercase text:value")
		  Assert.IsTrue(TestDictionary.Lookup(textKey2, 2) = "MixedCase Text:value")
		  Assert.IsTrue(TestDictionary.Lookup(textKey3, 3) = "UPPERCASE TEXT:value")
		  Assert.IsTrue(TestDictionary.Lookup(stringKey1, 4) = "lowercase string:value")
		  Assert.IsTrue(TestDictionary.Lookup(stringKey2, 5) = "MixedCase String:value")
		  Assert.IsTrue(TestDictionary.Lookup(stringKey3, 6) = "UPPERCASE STRING:value")
		  
		  // Test using literals.
		  Assert.IsTrue(TestDictionary.Lookup("lowercase text", 1) = "lowercase text:value")
		  Assert.IsTrue(TestDictionary.Lookup("MixedCase Text", 2) = "MixedCase Text:value")
		  Assert.IsTrue(TestDictionary.Lookup("UPPERCASE TEXT", 3) = "UPPERCASE TEXT:value")
		  Assert.IsTrue(TestDictionary.Lookup("lowercase string", 4) = "lowercase string:value")
		  Assert.IsTrue(TestDictionary.Lookup("MixedCase String", 5) = "MixedCase String:value")
		  Assert.IsTrue(TestDictionary.Lookup("UPPERCASE STRING", 6) = "UPPERCASE STRING:value")
		  
		  // Fails.
		  Assert.IsFalse(TestDictionary.Lookup("LOWERCASE TEXT", 1) = "lowercase text:value")
		  Assert.IsFalse(TestDictionary.Lookup("MIXEDCASE TEXT", 2) = "MixedCase Text:value")
		  Assert.IsFalse(TestDictionary.Lookup("uppercase text", 3) = "UPPERCASE TEXT:value")
		  Assert.IsFalse(TestDictionary.Lookup("UPPERCASE STRING", 4) = "lowercase string:value")
		  Assert.IsFalse(TestDictionary.Lookup("MIXEDCASE STRING", 5) = "MixedCase String:value")
		  Assert.IsFalse(TestDictionary.Lookup("lowercase string", 6) = "UPPERCASE STRING:value")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ValueAssignTest()
		  Var bd As New BetterDictionary(True)
		  
		  Var textKey1 As Text = "lowercase text"
		  Var textKey2 As Text = "MixedCase Text"
		  Var textKey3 As Text = "UPPERCASE TEXT"
		  Var stringKey1 As Text = "lowercase string"
		  Var stringKey2 As Text = "MixedCase String"
		  Var stringKey3 As Text = "UPPERCASE STRING"
		  
		  // Assign with variables.
		  bd.Value(textKey1) = "lowercase text:value"
		  bd.Value(textKey2) = "MixedCase Text:value"
		  bd.Value(textKey3) = "UPPERCASE TEXT:value"
		  bd.Value(stringKey1) = "lowercase string:value"
		  bd.Value(stringKey2) = "MixedCase String:value"
		  bd.Value(stringKey3) = "UPPERCASE STRING:value"
		  
		  Assert.IsTrue(bd.Value(textKey1) = "lowercase text:value")
		  Assert.IsTrue(bd.Value(textKey2) = "MixedCase Text:value")
		  Assert.IsTrue(bd.Value(textKey3) = "UPPERCASE TEXT:value")
		  Assert.IsTrue(bd.Value(stringKey1) = "lowercase string:value")
		  Assert.IsTrue(bd.Value(stringKey2) = "MixedCase String:value")
		  Assert.IsTrue(bd.Value(stringKey3) = "UPPERCASE STRING:value")
		  
		  // Assign with literals.
		  bd.Value("lowercase text") = "lowercase text:value"
		  bd.Value("MixedCase Text") = "MixedCase Text:value"
		  bd.Value("UPPERCASE TEXT") = "UPPERCASE TEXT:value"
		  bd.Value("lowercase string") = "lowercase string:value"
		  bd.Value("MixedCase String") = "MixedCase String:value"
		  bd.Value("UPPERCASE STRING") = "UPPERCASE STRING:value"
		  
		  Assert.IsTrue(bd.Value(textKey1) = "lowercase text:value")
		  Assert.IsTrue(bd.Value(textKey2) = "MixedCase Text:value")
		  Assert.IsTrue(bd.Value(textKey3) = "UPPERCASE TEXT:value")
		  Assert.IsTrue(bd.Value(stringKey1) = "lowercase string:value")
		  Assert.IsTrue(bd.Value(stringKey2) = "MixedCase String:value")
		  Assert.IsTrue(bd.Value(stringKey3) = "UPPERCASE STRING:value")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ValueRetrieveTest()
		  Var textKey1 As Text = "lowercase text"
		  Var textKey2 As Text = "MixedCase Text"
		  Var textKey3 As Text = "UPPERCASE TEXT"
		  Var stringKey1 As Text = "lowercase string"
		  Var stringKey2 As Text = "MixedCase String"
		  Var stringKey3 As Text = "UPPERCASE STRING"
		  
		  // Test with variables.
		  Assert.IsTrue(TestDictionary.Value(textKey1) = "lowercase text:value")
		  Assert.IsTrue(TestDictionary.Value(textKey2) = "MixedCase Text:value")
		  Assert.IsTrue(TestDictionary.Value(textKey3) = "UPPERCASE TEXT:value")
		  Assert.IsTrue(TestDictionary.Value(stringKey1) = "lowercase string:value")
		  Assert.IsTrue(TestDictionary.Value(stringKey2) = "MixedCase String:value")
		  Assert.IsTrue(TestDictionary.Value(stringKey3) = "UPPERCASE STRING:value")
		  
		  // Test with literals.
		  Assert.IsTrue(TestDictionary.Value("lowercase text") = "lowercase text:value")
		  Assert.IsTrue(TestDictionary.Value("MixedCase Text") = "MixedCase Text:value")
		  Assert.IsTrue(TestDictionary.Value("UPPERCASE TEXT") = "UPPERCASE TEXT:value")
		  Assert.IsTrue(TestDictionary.Value("lowercase string") = "lowercase string:value")
		  Assert.IsTrue(TestDictionary.Value("MixedCase String") = "MixedCase String:value")
		  Assert.IsTrue(TestDictionary.Value("UPPERCASE STRING") = "UPPERCASE STRING:value")
		  
		End Sub
	#tag EndMethod


	#tag ComputedProperty, Flags = &h21
		#tag Getter
			Get
			  Static bd As BetterDictionary = ConstructTestDictionary
			  Return bd
			  
			End Get
		#tag EndGetter
		Private TestDictionary As BetterDictionary
	#tag EndComputedProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
