Function get_date_order()

'/* licenced under GPL v 2


'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, see
'<https://www.gnu.org/licenses/>.


'author : manolo ariza manolo.ar.cu@gmail.com

' This functions returns format date.
' Order of date elements: 0 = month-day-year, 1 = day-month-year, 2 = year-month-day
' please note that this function is read only and comes from the OS
' if you need to change your date format,
' you need to change your locale settings and then restart excel. google is your friend



Dim x As Integer
Dim tmp As String

x = Application.International(32) 'get the value that determines the date format

Select Case x
Case 0
    tmp = "Your date format is MM/DD/YYYY"
Case 1
    tmp = "Your date format is DD/MM/YYYY"
Case 2
    tmp = "Your date format is YYYY/MM/DD"
End Select

get_date_order = tmp

End Function

