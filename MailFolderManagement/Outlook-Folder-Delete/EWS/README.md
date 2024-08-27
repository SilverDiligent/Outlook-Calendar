# dev-outlook-folder-delete

#################################################################################################################################
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. # 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  #
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.               #
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object  #
# code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software   #
# product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the  #
# Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims   #
# or lawsuits, including attorneysÃ¢â‚¬â„¢ fees, that arise or result from the use or distribution of the Sample Code.#
#################################################################################################################################
 
#----------------------------------------------------------------------              
#-     UPDATE VARIABLES TO REFLECT YOUR ENVIRONMENT                   -
#----------------------------------------------------------------------


# Download the EWS Managed API from here:
# https://github.com/officedev/ews-managed-api
# In order to install, make sure you have a NuGet PackageSource for location https://www.nuget.org/api/v2 installed
# Check running Get-PackageSource
# You will find a Nuget package Source for https://api.nuget.org/v3/index.json and you will have to add another one:
# Register-PackageSource -Name MyNuGet -Location https://www.nuget.org/api/v2 -ProviderName NuGet
# Now find the WebServices package using command:
# Find-Package Microsoft.Exchange.WebServices -RequiredVersion 2.2.0 -Source MyNuGet
# Then install it by piping the previous command to the install-package command
# Find-Package Microsoft.Exchange.WebServices -RequiredVersion 2.2.0 -Source MyNuGet | install-package -scopes "currentuser" -force
 