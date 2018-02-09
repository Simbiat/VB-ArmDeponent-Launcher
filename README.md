# VB-ArmDeponent-Launcher
This script was designed to elevate some headaches caused by using Kazakhstan's [ArmDeponent](http://www.kacd.kz/ru/software/) in our companie's environment.

We have 2 departments working with it and they have 2 different databases. We also need to be able to easily to switch to CoB (Reserve) server, if and when requried and check that workstation has all the requisites for launching ArmDeponent, as well as user has appropriate rights. We have a launcher (launchingapp.exe) that manages user rights in the database, and provides them in a `01010101` like line, where each number is a boolean value of an "entitlement" in the application.

When launching the script and passing such a line to it, we check if user has appropriate rights and they do not mix for security reasons and only then move forward. If for some reason user does has access to multiple tables we allow selection of a database. Note, that, since script is parsing a string, you may need to adjust it, if you use it.

After rights checks, we validate that there is an Oracle Client of an appropriate vesion, that we have correct language settings. If something is wrong - appropriate message is shown. If everything is fine - we genearte an .ini file for ArmDeponent with appropriate database settings.

Script allows "cob.txt" file as a flag, in case we need to easily switch to CoB server. In this case end-users do not need to keep the technical stuff in their heads.
