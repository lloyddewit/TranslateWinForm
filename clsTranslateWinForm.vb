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
Imports System.Data.SQLite
Imports System.Windows.Forms

'''------------------------------------------------------------------------------------------------
''' <summary>   
''' Provides utility functions to translate the text in WinForm objects (e.g. menu items, forms and 
''' controls) to a different natural language (e.g. to French). 
''' <para>
''' This class uses an SQLite database to translate text items to a new language. The database must contain the following tables:
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
    '''     TODO this function is still under development - please do not peer review or test yet. 
    '''     Attempts to translate all the text in <paramref name="clsForm"/> to <paramref name="strLanguage"/>.
    ''' </summary>
    '''
    ''' <param name="clsForm">          The WinForm form to translate. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguage">      The language code to translate to (e.g. 'fr' for French). </param>
    '''
    ''' <returns>   If an exception is thrown, then returns the exception text; else returns Nothing. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function translateForm(clsForm As Form, strDataSource As String, strLanguage As String) As String
        Try
            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                Using clsCommand As New SQLiteCommand(clsConnection)

                    'get all translations for the specified form and language
                    clsCommand.CommandText = "SELECT control_name, translation FROM form_controls, translations WHERE form_name = '" & clsForm.Name &
                                         "' AND language_code = '" & strLanguage & "' And form_controls.id_text = translations.id_text"
                    Dim clsReader As SQLiteDataReader = clsCommand.ExecuteReader()
                    Using clsReader

                        'for each translation row
                        While (clsReader.Read())

                            'translate the control's text to the new language
                            Dim strControlName As String = clsReader.GetString(0)
                            Dim strTranslation As String = clsReader.GetString(1)
                            CallByName(clsForm.Controls(strControlName), "Text", CallType.Set, strTranslation)

                        End While
                    End Using
                End Using
                clsConnection.Close()
            End Using
        Catch e As Exception
            Return e.Message & Environment.NewLine &
                    "A problem occured attempting to translate the form " &
                    If(IsNothing(clsForm), "(null value)", clsForm.Name) &
                    " to language " & strLanguage & " using database " & strDataSource & "."
        End Try
        Return Nothing
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Attempts to translate all the text in the menu items in <paramref name="tsCollection"/> 
    '''     to <paramref name="strLanguage"/>.
    ''' </summary>
    '''
    ''' <param name="tsCollection">     The WinForm menu items to translate. </param>
    ''' <param name="ctrParent">        The WinForm control that is the parent of the menu. </param>
    ''' <param name="strDataSource">    The path of the SQLite '.db' file that contains the
    '''                                 translation database. </param>
    ''' <param name="strLanguage">      (Optional) The language code to translate to (e.g. 'fr' for
    '''                                 French). </param>
    '''
    ''' <returns>
    '''     If an exception is thrown, then returns the exception text; else returns 'Nothing'.
    ''' </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function translateMenuItems(tsCollection As ToolStripItemCollection,
                                              ctrParent As Control, strDataSource As String,
                                              strLanguage As String) As String
        Try
            Dim dctMenuItems As Dictionary(Of String, ToolStripMenuItem) = GetDctMenuItems(tsCollection)

            'connect to the SQLite database that contains the translations
            Dim clsBuilder As New SQLiteConnectionStringBuilder With {
                .FailIfMissing = True,
                .DataSource = strDataSource}
            Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
                clsConnection.Open()
                Using clsCommand As New SQLiteCommand(clsConnection)

                    'get all translations for the specified form and language
                    clsCommand.CommandText = "SELECT control_name, translation FROM form_controls, translations WHERE form_name = '" & ctrParent.Name &
                                             "' AND language_code = '" & strLanguage & "' And form_controls.id_text = translations.id_text"
                    Using clsReader As SQLiteDataReader = clsCommand.ExecuteReader()

                        'for each translation row
                        While (clsReader.Read())

                            'ignore rows where the translation text is null or missing
                            If clsReader.FieldCount < 2 OrElse clsReader.IsDBNull(1) Then
                                Continue While
                            End If

                            'translate the menu item's text to the new language
                            Dim strMenuItemName As String = clsReader.GetString(0)
                            Dim strTranslation As String = clsReader.GetString(1)
                            Dim mnuItem As ToolStripMenuItem = Nothing
                            If dctMenuItems.TryGetValue(strMenuItemName, mnuItem) Then
                                mnuItem.Text = strTranslation
                            End If

                        End While
                    End Using
                End Using
                clsConnection.Close()
            End Using
        Catch e As Exception
            Return e.Message & Environment.NewLine &
                    "A problem occured attempting to translate the menu associated with form " &
                    If(IsNothing(ctrParent), "(null value)", ctrParent.Name) &
                    " to language " & strLanguage & " using database " & strDataSource & "."
        End Try
        Return Nothing
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    '''     Recursively traverses the <paramref name="tsCollection"/> menu hierarchy and returns a 
    '''     dictionary containing the name of each (sub)menu option in 
    '''     <paramref name="tsCollection"/> (as the dictionary key), together with its associated 
    '''     object (as the dictionary value). 
    ''' </summary>
    '''
    ''' <param name="tsCollection"> The WinForm menu item hierarchy used to populate the returned 
    '''                             dictionary. </param>
    '''
    ''' <returns>   
    '''     A dictionary containing the name of each (sub)menu option in 
    '''     <paramref name="tsCollection"/> (as the dictionary key), together with it's associated 
    '''     object (as the dictionary value). 
    ''' </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function GetDctMenuItems(tsCollection As ToolStripItemCollection) As Dictionary(Of String, ToolStripMenuItem)
        Dim dctMenuItems As Dictionary(Of String, ToolStripMenuItem) = New Dictionary(Of String, ToolStripMenuItem)

        For Each tsItem As ToolStripItem In tsCollection
            If Not String.IsNullOrEmpty(tsItem.Name) AndAlso Not String.IsNullOrEmpty(tsItem.Text) AndAlso
                    Not dctMenuItems.ContainsKey(tsItem.Name) Then
                dctMenuItems.Add(tsItem.Name, tsItem)
            End If
            Dim mnuItem As ToolStripMenuItem = TryCast(tsItem, ToolStripMenuItem)
            If mnuItem IsNot Nothing AndAlso mnuItem.HasDropDownItems Then
                Dim dctSubMenuItems As Dictionary(Of String, ToolStripMenuItem) = GetDctMenuItems(mnuItem.DropDownItems)
                dctMenuItems = dctMenuItems.Union(dctSubMenuItems).ToDictionary(Function(p) p.Key, Function(p) p.Value)
            End If
        Next

        Return dctMenuItems
    End Function

End Class
