Module GestioneXML
    Private XMLDoc_Settings As XDocument
    Private XMLDoc_Strutture As XDocument
    Private XMLDoc_Language As XDocument
    Public Event UpdateLivePreviewEvent As EventHandler 'ugly hack to make live preview, i don't know a solution to check obscoll items changed

#Region "Setting.xml"

    Public Sub XMLSettings_LoadFile()
        Try
            XMLDoc_Settings = XDocument.Load(DataPath & "\settings.xml")

        Catch ex As Exception
            Throw New Exception("Read error: settings.xml", ex)
        End Try
    End Sub

    Public Sub XMLSettings_WriteFile()
        Try
            XMLDoc_Settings.Save(DataPath & "\settings.xml", SaveOptions.None)

        Catch ex As Exception
            Throw New Exception("Write error: settings.xml", ex)
        End Try
    End Sub

    Public Function XMLSettings_Read(ByVal Chiave As String) As String
        Return XMLDoc_Settings.<Impostazioni>.First.Element(Chiave).Value
    End Function

    Public Sub XMLSettings_Save(ByVal Chiave As String, ByRef Value As String)
        XMLDoc_Settings.<Impostazioni>.First.Element(Chiave).Value = Value
    End Sub

#End Region

#Region "Strutture.xml"

    Public Sub XMLStrutture_LoadFile()
        Try
            XMLDoc_Strutture = XDocument.Load(DataPath & "\strutture_categorie.xml")

        Catch ex As Exception
            Throw New Exception("Read error: strutture_categorie.xml", ex)
        End Try
    End Sub

    Public Function XMLStrutture_GetXML() As String
        Dim ElementToCopy As New XElement(XMLSettings_Read("CategoriaSelezionata"))
        ElementToCopy.SetAttributeValue("nome", "Clipboard")

        For Each item In Coll_Struttura
            Dim NuovoElemento As XElement
            If Not String.IsNullOrEmpty(item.Opzioni.Trim) Then
                NuovoElemento = XElement.Parse(item.Opzioni)
            Else
                NuovoElemento = XElement.Parse("<" & item.TipoDato & " />")
            End If
            ElementToCopy.Add(NuovoElemento)
        Next

        Return ElementToCopy.ToString
    End Function

    Public Sub XMLStrutture_SetXML(_ClipBoardXML As String)
        Dim ElementToParse As XElement = Nothing
        Try
            ElementToParse = XElement.Parse(_ClipBoardXML.Trim)
        Catch ex As Exception
            Throw New Exception("Impossibile incollare il preset, ci sono errori nel codice.")
        End Try

        If ElementToParse IsNot Nothing Then

            'TODO: verificare stringhe tipodati

            If Not ElementToParse.Name.ToString.Equals(XMLSettings_Read("CategoriaSelezionata")) Then
                Throw New Exception("Impossibile incollare il preset, categoria sbagliata.")
            End If
            If Not ElementToParse.Attribute("nome").Value.Equals("ClipBoard") Then
                Throw New Exception("Impossibile incollare, ci sono errori nel codice.")
            End If

            Coll_Struttura.Clear()

            For Each child In ElementToParse.Elements
                'This function (in the comment) return the xml code of child element without root xml element
                'Debug.Print(String.Join("", child.Nodes.[Select](Function(ele) ele.ToString)))
                Dim Contenuto As String = child.ToString
                AggiungiDatoStrutturaCollection(child.Name.ToString, Contenuto)
            Next
        End If
    End Sub

    Private Sub XMLStrutture_WriteFile()
        Try
            XMLDoc_Strutture.Save(DataPath & "\strutture_categorie.xml", SaveOptions.None)

        Catch ex As Exception
            Throw New Exception("Write error: strutture_categorie.xml", ex)
        End Try
    End Sub

    Public Function XMLStrutture_Elenca() As ArrayList
        Dim Elenco As New ArrayList
        For Each element In XMLDoc_Strutture.<strutture>.Elements(XMLSettings_Read("CategoriaSelezionata"))
            If element.Attribute("nome").Value = "Default" Then Continue For
            Elenco.Add(element.Attribute("nome").Value)
        Next
        Return Elenco
    End Function

    Public Sub XMLStrutture_Read(ByVal NomeStruttura As String)
        If NomeStruttura Is Nothing Then NomeStruttura = "Default"

        Coll_Struttura.Clear()

        For Each element In XMLDoc_Strutture.<strutture>.Elements(XMLSettings_Read("CategoriaSelezionata"))
            If element.Attribute("nome").Value = NomeStruttura Then

                For Each child In element.Elements
                    'This function (in the comment) return the xml code of child element without root xml element
                    'Debug.Print(String.Join("", child.Nodes.[Select](Function(ele) ele.ToString)))
                    Dim Contenuto As String = child.ToString
                    AggiungiDatoStrutturaCollection(child.Name.ToString, Contenuto)
                Next
            End If
        Next
    End Sub

    Public Sub XMLStrutture_Save(ByVal NomeStruttura As String)
        Dim ElementToUpdate As New XElement(XMLSettings_Read("CategoriaSelezionata"))
        ElementToUpdate.SetAttributeValue("nome", NomeStruttura)
        Dim InsertNew As Boolean = True

        For Each element In XMLDoc_Strutture.<strutture>.Elements(XMLSettings_Read("CategoriaSelezionata"))
            If element.Attribute("nome").Value = NomeStruttura Then
                ElementToUpdate = element
                InsertNew = False
                Exit For
            End If
        Next

        ElementToUpdate.RemoveNodes()

        For Each item In Coll_Struttura
            Dim NuovoElemento As XElement
            If Not String.IsNullOrEmpty(item.Opzioni.Trim) Then
                NuovoElemento = XElement.Parse(item.Opzioni)
            Else
                NuovoElemento = XElement.Parse("<" & item.TipoDato & " />")
            End If
            ElementToUpdate.Add(NuovoElemento)
        Next

        If InsertNew Then XMLDoc_Strutture.Descendants("strutture").Last.Add(ElementToUpdate)

        XMLStrutture_WriteFile()
    End Sub

    Public Sub XMLStrutture_Delete(ByVal NomeStruttura As String)

        For Each element In XMLDoc_Strutture.<strutture>.Elements(XMLSettings_Read("CategoriaSelezionata"))
            If element.Attribute("nome").Value = NomeStruttura Then
                element.Remove()
                Exit For
            End If
        Next

        XMLStrutture_WriteFile()
    End Sub

    Public Function LeggiStringOpzioni(ByVal NomeDato As String, ByVal DatiXML As String, Optional OverrideFallBack As String = Nothing) As String

        If String.IsNullOrEmpty(DatiXML) Then
            Return IIf(OverrideFallBack Is Nothing, FallbackDefaultDataStruttura.Item(NomeDato), OverrideFallBack)
        Else

            Dim RootXDoc = XDocument.Parse(DatiXML, LoadOptions.PreserveWhitespace)
            Dim valore = RootXDoc.Elements.First().Elements(NomeDato).Value

            If valore Is Nothing Then
                Return IIf(OverrideFallBack Is Nothing, FallbackDefaultDataStruttura.Item(NomeDato), OverrideFallBack)
            Else
                Debug.Print("LeggiStringOpzioni: Nomedato: " & NomeDato & " |Nuovo valore: '" & valore & "'")

                Return valore
            End If
        End If

    End Function

    Public Function ModificaStringOpzione(ByVal NomeDato As String, ByVal NuovoValore As String, ByRef itemSel As ItemStruttura)
        Debug.Print("ModificaStringOpzione: Nomedato: " & NomeDato & " |Nuovo valore: '" & NuovoValore & "'")

        Dim ElementRoot As XElement
        If String.IsNullOrEmpty(itemSel.Opzioni) Then

            ElementRoot = New XElement(itemSel.TipoDato)
            Dim ChildElement = New XElement(NomeDato, NuovoValore)
            ElementRoot.Add(ChildElement)
            'Debug.Print("ModStringOpz: Aggiunto")
        Else

            ElementRoot = XElement.Parse(itemSel.Opzioni)
            Dim ElementToUpdate = Nothing

            If ElementRoot.HasElements Then
                If ElementRoot.Elements(NomeDato).Any Then
                    For Each node As XElement In ElementRoot.Elements
                        If node.Name.ToString.Equals(NomeDato) Then
                            ElementToUpdate = node
                            'Debug.Print("ModStringOpz: Nodo trovato")
                            Exit For
                        End If
                    Next
                End If
            End If

            If ElementToUpdate Is Nothing Then
                Dim ChildElement = New XElement(NomeDato, NuovoValore)
                ElementRoot.Add(ChildElement)
                'Debug.Print("ModStringOpz: Aggiunto solo child")
            Else

                ElementToUpdate.Value = NuovoValore
                'Debug.Print("ModStringOpz: Trovato e modificato")
            End If
        End If
        'Debug.Print("ModStringOpz: Result" & vbCrLf & ElementRoot.ToString)

        RaiseEvent UpdateLivePreviewEvent(Nothing, Nothing)

        Return ElementRoot.ToString
    End Function

#End Region



#Region "language_?.xml"
    'NOW ALL INCLUDED IN RESOURCE FILES

    'Public Sub XMLLanguage_LoadFile()
    'Try
    '        XMLDoc_Language = XDocument.Load(DataPath & "\language_IT.xml")
    '
    'Catch ex As Exception
    ' Throw New Exception("Read error: language_?.xml", ex)
    ' End Try
    ' End Sub
    '
    '    Public Sub XMLLanguage_WriteFile()
    '    Try
    '            XMLDoc_Language.Save(DataPath & "\language_IT.xml", SaveOptions.None)
    '
    '    Catch ex As Exception
    '    Throw New Exception("Write error: language_?.xml", ex)
    '    End Try
    '    End Sub

    ' Public Function XMLLanguage_Read(ByVal Chiave As String) As String
    '  Return XMLDoc_Language.<languages>.First.Element(Chiave).Value
    '  End Function

    ''for now not needed
    ''    Public Sub XMLLanguage_Save(ByVal Chiave As String, ByRef Value As String)
    ''        XMLDoc_Language.<languages>.First.Element(Chiave).Value = Value
    ''    End Sub

#End Region
End Module
