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

Public Class clsTranslateWinForm
    Public Shared Sub translateForm(clsForm As Form, strDataSource As String, Optional strLanguage As String = "")
        'connect to the SQLite database that contains the translations
        Dim clsBuilder As New SQLiteConnectionStringBuilder With {
            .FailIfMissing = True,
            .DataSource = strDataSource
        }
        Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
            clsConnection.Open()
            Using clsCommand As New SQLiteCommand(clsConnection)

                'get all translations for the specified form and language
                clsCommand.CommandText = "SELECT control_name, translation FROM form_controls, translations WHERE form_name = """ & clsForm.Name &
                                         """ AND language_code = """ & strLanguage & """ And form_controls.id_text = translations.id_text"
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
    End Sub

    Public Shared Sub translateMenuItems(tsCollection As ToolStripItemCollection, ctrParent As Control, strDataSource As String, Optional strLanguage As String = "")

        Dim dctMenuItems As Dictionary(Of String, ToolStripMenuItem) = GetDctMenuItems(tsCollection)

        'connect to the SQLite database that contains the translations
        Dim clsBuilder As New SQLiteConnectionStringBuilder With {
            .FailIfMissing = True,
            .DataSource = strDataSource
        }
        Using clsConnection As New SQLiteConnection(clsBuilder.ConnectionString)
            clsConnection.Open()
            Using clsCommand As New SQLiteCommand(clsConnection)

                'get all translations for the specified form and language
                clsCommand.CommandText = "SELECT control_name, translation FROM form_controls, translations WHERE form_name = """ & ctrParent.Name &
                                         """ AND language_code = """ & strLanguage & """ And form_controls.id_text = translations.id_text"
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
    End Sub

    Private Shared Function GetDctMenuItems(tsCollection As ToolStripItemCollection) As Dictionary(Of String, ToolStripMenuItem)
        Dim dctMenuItems As Dictionary(Of String, ToolStripMenuItem) = New Dictionary(Of String, ToolStripMenuItem)

        For Each tsItem As ToolStripItem In tsCollection
            If Not String.IsNullOrEmpty(tsItem.Text) Then
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
