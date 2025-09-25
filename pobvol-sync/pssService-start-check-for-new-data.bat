@echo off

#
# Script:	    pssService-start-check-for-new-data.bat
# Task:		    Process new and changed data
# 
#This file is part of the software solution pobvol Service Solution. 
#pobvol Service Solution is Free Software, delivered as open source. You can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. The solution is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with the solution. If not, see <http://www.gnu.org/licenses/>. 
#Copyright © 2025 Volker Pobloth
#Web: https://pobvol.com/
#
#---------------------------------------------------------------------------------------

set location=%~dp0
cd %location%
cmd /c start /min "" pwsh -noprofile -ExecutionPolicy Bypass -WindowStyle Hidden -WorkingDirectory %location% -File "pssService-check-for-new-data.ps1"

