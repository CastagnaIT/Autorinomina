Imports System.IO
Imports System.Xml.Serialization
''' <summary>
''' Functions for performing common XML Serialization operations.
''' <para>Only public properties and variables will be serialized.</para>
''' <para>Use the [XmlIgnore] attribute to prevent a property/variable from being serialized.</para>
''' <para>Object to be serialized must have a parameterless constructor.</para>
''' </summary>
Public NotInheritable Class XmlSerialization
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Writes the given object instance to an XML file.
    ''' <para>Only Public properties and variables will be written to the file. These can be any type though, even other classes.</para>
    ''' <para>If there are public properties/variables that you do not want written to the file, decorate them with the [XmlIgnore] attribute.</para>
    ''' <para>Object type must have a parameterless constructor.</para>
    ''' </summary>
    ''' <typeparam name="T">The type of object being written to the file.</typeparam>
    ''' <param name="filePath">The file path to write the object instance to.</param>
    ''' <param name="objectToWrite">The object instance to write to the file.</param>
    ''' <param name="append">If false the file will be overwritten if it already exists. If true the contents will be appended to the file.</param>
    Public Shared Sub WriteToXmlFile(Of T As New)(filePath As String, objectToWrite As T, RootAttributeName As String, Optional append As Boolean = False)
        Dim writer As TextWriter = Nothing
        Try
            Dim serializer = New XmlSerializer(GetType(T), New XmlRootAttribute(RootAttributeName))
            writer = New StreamWriter(filePath, append)
            serializer.Serialize(writer, objectToWrite)
        Finally
            If writer IsNot Nothing Then
                writer.Close()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Reads an object instance from an XML file.
    ''' <para>Object type must have a parameterless constructor.</para>
    ''' </summary>
    ''' <typeparam name="T">The type of object to read from the file.</typeparam>
    ''' <param name="filePath">The file path to read the object instance from.</param>
    ''' <returns>Returns a new instance of the object read from the XML file.</returns>
    Public Shared Function ReadFromXmlFile(Of T As New)(filePath As String, RootAttributeName As String) As List(Of T)
        Dim reader As TextReader = Nothing
        Try
            Dim serializer = New XmlSerializer(GetType(List(Of T)), New XmlRootAttribute(RootAttributeName))
            reader = New StreamReader(filePath)
            Return DirectCast(serializer.Deserialize(reader), List(Of T))
        Finally
            If reader IsNot Nothing Then
                reader.Close()
            End If
        End Try
    End Function
End Class
