' IDEMS International
' Copyright (C) 2021
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License 
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
Imports System.ComponentModel
Imports System.Data.SQLite
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

'''------------------------------------------------------------------------------------------------
''' <summary>   
''' Provides utility functions to translate the text in WinForm objects (e.g. menu items, forms and 
''' controls) to a different natural language (e.g. to French). 
''' <para>
''' This class uses an SQLite database to translate text items to a new language. The database must 
''' contain the following tables:
''' <code>
''' CREATE TABLE "form_controls" (
'''	"form_name"	TEXT,
'''	"control_name"	TEXT,
'''	"id_text"	TEXT,
'''	PRIMARY KEY("form_name", "control_name")
''' )
''' <para>
''' CREATE TABLE "translations" (
'''	"id_text"	TEXT,
'''	"language_code"	TEXT,
'''	"translation"	TEXT,
'''	PRIMARY KEY("id_text", "language_code")
''' )
''' </para></code>
''' For example, if the 'form_controls' table contains a row with the values 
''' {'frmMain', 'mnuFile', 'File'}, 
''' then the 'translations' table should have a row for each supported language, e.g. 
''' {'File', 'en', 'File'}, {'File', 'fr', 'Fichier'}.
''' </para><para>
''' Note: This class is intended to be used solely as a 'static' class (i.e. contains only shared 
''' members, cannot be instantiated and cannot be inherited from).
''' In order to enforce this (and prevent developers from using this class in an unintended way), 
''' the class is declared as 'NotInheritable` and the constructor is declared as 'Private'.</para>
''' </summary>
'''------------------------------------------------------------------------------------------------

Public NotInheritable Class clsTranslateWinForm

    '''--------------------------------------------------------------------------------------------
    ''' <summary> 
    ''' Declare constructor 'Private' to prevent instantiation of this class (see class comments 
    ''' for more details). 
    ''' </summary>
    '''--------------------------------------------------------------------------------------------
    Private Sub New()
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Translates all the text in form <paramref name="clsForm"/> into language 
    '''     <paramref name="strLanguage"/> using the translations in database 
    '''     <paramref name="strDataSource"/>.
    '''     All the form's (sub)controls and (sub) menus are translated.     
    ''' </summary>
    '''
    ''' <param name="clsForm">          The WinForm form to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French). 
    '''                                 </param>
    '''
    ''' <returns>   If an exception is thrown, then returns the exception text; else returns 
    '''             'Nothing'. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function TranslateForm(clsForm As Form, strDataSource As String,
                                         strLanguage As String) As String
        If IsNothing(clsForm) OrElse String.IsNullOrEmpty(strDataSource) OrElse
                String.IsNullOrEmpty(strLanguage) Then
            Return ("Developer Error: Illegal parameter passed to TranslateForm (language: " &
                   strLanguage & ", source: " & strDataSource & ").")
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        GetDctComponentsFromControl(clsForm, dctComponents)
        Return TranslateDctComponents(dctComponents, clsForm.Name, strDataSource, strLanguage)

    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Translates all the (sub)menu items in <paramref name="clsMenuItems"/> into language
    '''    <paramref name="strLanguage"/> using the translations in database
    '''    <paramref name="strDataSource"/>.
    ''' </summary>
    '''
    ''' <param name="strParentName">    The menu's parent control. </param>
    ''' <param name="clsMenuItems">     The (sub)menu items to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French).
    '''                                 </param>
    '''
    ''' <returns>   If an exception is thrown, then returns the exception text; else returns 
    '''             'Nothing'. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function TranslateMenuItems(strParentName As String, clsMenuItems As ToolStripItemCollection,
                                              strDataSource As String, strLanguage As String) As String
        If IsNothing(clsMenuItems) OrElse String.IsNullOrEmpty(strParentName) OrElse
                String.IsNullOrEmpty(strDataSource) OrElse String.IsNullOrEmpty(strLanguage) Then
            Return ("Developer Error: Illegal parameter passed to TranslateMenuItems (language: " &
                   strLanguage & ", source: " & strDataSource & ", parent: " & strParentName & " ).")
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        GetDctComponentsFromMenuItems(clsMenuItems, dctComponents)

        Return TranslateDctComponents(dctComponents, strParentName, strDataSource, strLanguage)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Returns <paramref name="strText"/> translated into <paramref name="strLanguage"/>. 
    '''    <para>
    '''    Translations can be bi-directional (e.g. from English to French or from French to English).
    '''    If <paramref name="strText"/> is already in the current language, or if no translation 
    '''    can be found, then returns <paramref name="strText"/>.         
    '''    </para></summary>
    '''
    ''' <param name="strText">          The text to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French).
    '''                                 </param>
    '''
    ''' <returns>   <paramref name="strText"/> translated into <paramref name="strLanguage"/>. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetTranslation(strText As String, strDataSource As String,
                                          strLanguage As String) As String
        Dim strTranslation As String = ""
        Try
            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                strTranslation = GetDynamicTranslation(strText, strLanguage, clsConnection)
                clsConnection.Close()
            End Using
        Catch e As Exception
            Return e.Message & Environment.NewLine &
                    "A problem occured attempting to translate string '" & strText &
                    "' to language " & strLanguage & " using database " & strDataSource & "."
        End Try
        Return strTranslation
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''     Recursively traverses the <paramref name="clsControl"/> control hierarchy and returns a
    '''     string containing the parent, name and associated text of each control. The string is 
    '''     formatted as a comma-separated list suitable for importing into a database.
    ''' </summary>
    '''
    ''' <param name="clsControl">   The control to process (it's children and sub-children shall 
    '''                             also be processed recursively). </param>
    '''
    ''' <returns>   
    '''     A string containing the parent, name and associated text of each control in the 
    '''     hierarchy. The string is formatted as a comma-separated list suitable for importing 
    '''     into a database. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetControlsAsCsv(clsControl As Control) As String
        If IsNothing(clsControl) Then
            Return ""
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        GetDctComponentsFromControl(clsControl, dctComponents)

        Dim strControlsAsCsv As String = ""
        For Each clsComponent In dctComponents
            If TypeOf clsComponent.Value Is Control Then
                Dim clsTmpControl As Control = DirectCast(clsComponent.Value, Control)
                strControlsAsCsv &= clsControl.Name & "," & clsComponent.Key & "," & GetCsvText(clsTmpControl.Text) & vbCrLf
            ElseIf TypeOf clsComponent.Value Is ToolStripItem Then
                Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent.Value, ToolStripItem)
                strControlsAsCsv &= clsControl.Name & "," & clsComponent.Key & "," & GetCsvText(clsMenuItem.Text) & vbCrLf
            Else
                MsgBox("Developer Error: Translation dictionary entry (" & clsControl.Name & "," & clsComponent.Key & ") contained unexpected value type.")
            End If
        Next

        Return strControlsAsCsv
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''     Recursively traverses the <paramref name="clsMenuItems"/> menu hierarchy and returns a 
    '''     string containing the parent, name and associated text of each (sub)menu option in 
    '''     <paramref name="clsMenuItems"/>. The string is formatted as a comma-separated list 
    '''     suitable for importing into a database.
    ''' </summary>
    '''
    ''' <param name="clsControl">        The WinForm control that is the parent of the menu. </param>
    ''' <param name="clsMenuItems">     The WinForm menu items to add to the return string. </param>
    '''
    ''' <returns>   
    '''     A string containing the parent and name of each (sub)menu option in
    '''     <paramref name="clsMenuItems"/>. The string is formatted as a comma-separated list
    '''     suitable for importing into a database. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function GetMenuItemsAsCsv(clsControl As Control, clsMenuItems As ToolStripItemCollection) As String
        If IsNothing(clsControl) OrElse IsNothing(clsMenuItems) Then
            Return ""
        End If

        Dim dctComponents As Dictionary(Of String, Component) = New Dictionary(Of String, Component)
        GetDctComponentsFromMenuItems(clsMenuItems, dctComponents)

        Dim strMenuItemsAsCsv As String = ""
        For Each clsComponent In dctComponents

            If TypeOf clsComponent.Value Is ToolStripItem Then
                Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent.Value, ToolStripItem)
                strMenuItemsAsCsv &= clsControl.Name & "," & clsComponent.Key & "," & GetCsvText(clsMenuItem.Text) & vbCrLf
            Else
                MsgBox("Developer Error: Translation dictionary entry (" & clsControl.Name & "," & clsComponent.Key & ") contained unexpected value type.")
            End If

        Next
        Return strMenuItemsAsCsv
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''    Populates dictionary <paramref name="dctComponents"/> with the control 
    '''    <paramref name="clsControl"/> and its children.    
    '''    The dictionary can then be used to conveniently translate the menu item text (see other
    '''    functions and subs in this class).
    ''' </summary>
    '''
    ''' <param name="clsControl">       The control used to populate the dictionary. </param>
    ''' <param name="dctComponents">    [in,out] Dictionary to store the control and its children. 
    '''                                 </param>
    '''--------------------------------------------------------------------------------------------
    Private Shared Sub GetDctComponentsFromControl(clsControl As Control, ByRef dctComponents As Dictionary(Of String, Component))
        If IsNothing(clsControl) OrElse IsNothing(clsControl.Controls) OrElse IsNothing(dctComponents) Then
            Exit Sub
        End If

        'if control is valid, then add it to the dictionary
        'If Not (String.IsNullOrEmpty(clsControl.Name) OrElse
        '            String.IsNullOrEmpty(clsControl.Text) OrElse
        '            clsControl.Text.Contains(vbCr) OrElse clsControl.Text.Contains(vbLf) OrElse 'ignore multiline text
        '            Not Regex.IsMatch(clsControl.Text, "[a-zA-Z]") OrElse 'ignore text that doesn't contain any letters (e.g. number strings)
        '            dctComponents.ContainsKey(clsControl.Name)) Then 'ignore components that are already in the dictionary
        If Not (String.IsNullOrEmpty(clsControl.Name) OrElse
                dctComponents.ContainsKey(clsControl.Name)) Then 'ignore components that are already in the dictionary
            dctComponents.Add(clsControl.Name, clsControl)
        End If

        For Each ctlChild As Control In clsControl.Controls

            'Recursively process different types of menus and child controls
            If TypeOf ctlChild Is MenuStrip Then
                Dim clsMenuStrip As MenuStrip = DirectCast(ctlChild, MenuStrip)
                GetDctComponentsFromMenuItems(clsMenuStrip.Items, dctComponents)
            ElseIf TypeOf ctlChild Is ToolStrip Then
                Dim clsToolStrip As ToolStrip = DirectCast(ctlChild, ToolStrip)
                GetDctComponentsFromMenuItems(clsToolStrip.Items, dctComponents)
            ElseIf TypeOf ctlChild Is Control Then
                GetDctComponentsFromControl(ctlChild, dctComponents)
            End If

        Next
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Populates dictionary <paramref name="dctComponents"/> with all the menu items, and 
    '''    sub-menu items in the <paramref name="clsMenuItems"/>. 
    '''    The dictionary can then be used to conveniently translate the menu item text (see other 
    '''    functions and subs in this class).
    ''' </summary>
    '''
    ''' <param name="clsMenuItems">     The list of menu items to populate the dictionary. </param>
    ''' <param name="dctComponents">    [in,out] Dictionary to store the menu items. </param>
    '''--------------------------------------------------------------------------------------------
    Private Shared Sub GetDctComponentsFromMenuItems(clsMenuItems As ToolStripItemCollection, ByRef dctComponents As Dictionary(Of String, Component))
        If IsNothing(clsMenuItems) OrElse IsNothing(dctComponents) Then
            Exit Sub
        End If

        For Each clsMenuItem As ToolStripItem In clsMenuItems

            'if menu item is valid, then add it to the dictionary
            If Not (String.IsNullOrEmpty(clsMenuItem.Name) OrElse
                    dctComponents.ContainsKey(clsMenuItem.Name)) Then 'ignore components that are already in the dictionary
                dctComponents.Add(clsMenuItem.Name, clsMenuItem)
            End If

            'Recursively process different types of sub-menu
            If TypeOf clsMenuItem Is ToolStripMenuItem Then
                Dim clsTmpMenuItem As ToolStripMenuItem = DirectCast(clsMenuItem, ToolStripMenuItem)
                If clsTmpMenuItem.HasDropDownItems Then
                    GetDctComponentsFromMenuItems(clsTmpMenuItem.DropDownItems, dctComponents)
                End If
            ElseIf TypeOf clsMenuItem Is ToolStripSplitButton Then
                Dim clsTmpMenuItem As ToolStripSplitButton = DirectCast(clsMenuItem, ToolStripSplitButton)
                If clsTmpMenuItem.HasDropDownItems Then
                    GetDctComponentsFromMenuItems(clsTmpMenuItem.DropDownItems, dctComponents)
                End If
            ElseIf TypeOf clsMenuItem Is ToolStripDropDownButton Then
                Dim clsTmpMenuItem As ToolStripDropDownButton = DirectCast(clsMenuItem, ToolStripDropDownButton)
                If clsTmpMenuItem.HasDropDownItems Then
                    GetDctComponentsFromMenuItems(clsTmpMenuItem.DropDownItems, dctComponents)
                End If
            End If

        Next
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Attempts to translate all the text in <paramref name="dctComponents"/>
    '''     to <paramref name="strLanguage"/>.
    '''     Opens database <paramref name="strDataSource"/> and reads in all translations for the 
    '''     <paramref name="strControlName"/> control for target language <paramref name="strLanguage"/>.
    '''     For each translation in the database, attempts to find the corresponding control or menu 
    '''     item in <paramref name="dctComponents"/>. If found, then it translates the text to the target language.
    ''' </summary>
    '''
    ''' <param name="dctComponents">    [in,out] The dictionary of translatable components. </param>
    ''' <param name="strControlName">   The name of the form or menu used to populate the dictionary. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French). </param>
    '''
    ''' <returns>
    '''     If an exception is thrown, then returns the exception text; else returns 'Nothing'.
    ''' </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function TranslateDctComponents(ByRef dctComponents As Dictionary(Of String, Component),
                                                   strControlName As String, strDataSource As String,
                                                   strLanguage As String) As String
        Try
            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                Using clsCommand As New SQLiteCommand(clsConnection)

                    'get all translations for the specified form and language
                    clsCommand.CommandText = "SELECT control_name, form_controls.id_text, translation FROM form_controls, translations WHERE form_name = '" & strControlName &
                                         "' AND language_code = '" & strLanguage & "' And form_controls.id_text = translations.id_text"
                    Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
                    Using clsReader

                        'for each translation row
                        While (clsReader.Read())

                            'ignore rows where the translation text is null or missing
                            If clsReader.FieldCount < 3 OrElse clsReader.IsDBNull(2) Then
                                Continue While
                            End If

                            'translate the component's text to the new language
                            Dim strComponentlName As String = clsReader.GetString(0)
                            Dim strIdText As String = clsReader.GetString(1)
                            Dim strTranslation As String = clsReader.GetString(2)
                            Dim clsComponent As Component = Nothing
                            If dctComponents.TryGetValue(strComponentlName, clsComponent) Then
                                If TypeOf clsComponent Is Control Then
                                    Dim clsControl As Control = DirectCast(clsComponent, Control)
                                    If strIdText = "ReplaceWithDynamicTranslation" Then
                                        clsControl.Text = GetDynamicTranslation(clsControl.Text, strLanguage, clsConnection)
                                    Else
                                        clsControl.Text = strTranslation
                                    End If
                                ElseIf TypeOf clsComponent Is ToolStripItem Then
                                    Dim clsMenuItem As ToolStripItem = DirectCast(clsComponent, ToolStripItem)
                                    clsMenuItem.Text = strTranslation
                                Else
                                    MsgBox("Developer Error: Translation dictionary entry (" & strComponentlName & ") contained unexpected value type.")
                                End If
                            End If

                        End While
                    End Using
                End Using
                clsConnection.Close()
            End Using
        Catch e As Exception
            Return e.Message & Environment.NewLine &
                    "A problem occured attempting to translate to language " & strLanguage &
                    " using database " & strDataSource & "."
        End Try
        Return Nothing
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Returns <paramref name="strText"/> translated into <paramref name="strLanguage"/>. 
    '''    <para>
    '''    Translations can be bi-directional (e.g. from English to French or from French to English).
    '''    If <paramref name="strText"/> is already in the current language, or if no translation 
    '''    can be found, then returns <paramref name="strText"/>.         
    '''    </para></summary>
    '''
    ''' <param name="strText">          The text to translate. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French).
    '''                                 </param>
    ''' <param name="clsConnection">    An open connection to the SQLite translation database. </param>
    '''
    ''' <returns>   <paramref name="strText"/> translated into <paramref name="strLanguage"/>. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function GetDynamicTranslation(strText As String, strLanguage As String, clsConnection As SQLiteConnection) As String
        If String.IsNullOrEmpty(strText) Then
            Return ""
        End If

        Using clsCommand As New SQLiteCommand(clsConnection)

            'in the translation text, convert any single quotes to make them suitable for the SQL command
            strText = strText.Replace("'", "''")

            'get all translations for the specified form and language
            'Note: The second `SELECT` is needed because we may sometimes need to translate  
            '      translated text back to the original text (e.g. from French to English when 
            '      the dialog language toggle button is clicked).
            clsCommand.CommandText = "SELECT translation FROM translations WHERE language_code = '" &
                                     strLanguage & "' AND id_text = '" & strText & "' OR (language_code = '" &
                                     strLanguage & "' AND id_text = " &
                                     "(SELECT id_text FROM translations WHERE translation = '" & strText & "'))"
            Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
            Using clsReader
                'for each translation row
                While (clsReader.Read())
                    'ignore rows where the translation text is null or missing
                    If clsReader.FieldCount < 1 OrElse clsReader.IsDBNull(0) Then
                        Continue While
                    End If
                    'return the translation text
                    Return clsReader.GetString(0)
                End While
            End Using
        End Using
        'if no tranlsation text was found then return original text unchanged
        Return strText
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''    Decides whether <paramref name="strText"/> is likely to be changed during execution of 
    '''    the software. If no, then returns <paramref name="strText"/>. If yes, then returns 
    '''    'ReplaceWithDynamicTranslation'. It makes the decision based upon a set of heuristics.
    '''    <para>
    '''    This function is normally only used when creating a comma-separated list suitable for 
    '''    importing into a database. During program execution, the 'ReplaceWithDynamicTranslation'
    '''    text tells the library to dynamically try and translate the current text, rather than
    '''    looking up the static text associated with the control.</para></summary>
    '''
    ''' <param name="strText">  The text to assess. </param>
    '''
    ''' <returns>   Decides whether <paramref name="strText"/> is likely to be changed during 
    '''             execution of the software. If no, then returns <paramref name="strText"/>. 
    '''             If yes, then returns'ReplaceWithDynamicTranslation'. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function GetCsvText(strText As String) As String
        If String.IsNullOrEmpty(strText) OrElse
                strText.Contains(vbCr) OrElse strText.Contains(vbLf) OrElse 'multiline text
                Regex.IsMatch(strText, "CheckBox\d+$") OrElse 'CheckBox1, CheckBox2 etc. normally indicates dynamic translation
                Regex.IsMatch(strText, "Label\d+$") OrElse 'Label1, Label2 etc. normally indicates dynamic translation
                Not Regex.IsMatch(strText, "[a-zA-Z]") Then 'text that doesn't contain any letters (e.g. number strings)
            Return "ReplaceWithDynamicTranslation"
        End If
        Return strText
    End Function


End Class
