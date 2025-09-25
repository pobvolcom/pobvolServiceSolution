//pssServiceBoard Startscreen TimerReadEinstellungen.fx
//Power Apps Language: en-US
//------------------------------------------------

Set(CurrentTime,Now());
//---------------------------------------------
If(IsBlank(CurrentYear),
    Set(CurrentYear,Year(Today()))
);
//---------------------------------------------
If(Connection.Connected,
    Set (OnlineStatus, 1),
// Else    
    Set (OnlineStatus, 0)
);
//---------------------------------------------
//Systemeinstellungen laden
If(!SystemSettingsLoaded || CountRows(colSystemSettings)=0,
    Set (_ShowLoadingHint,true);
    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="Einstellungen").Wert
    );

    //License
    Set(_userDomainCheck,true);
    Set(Lizenz,"valid");
    Set(MaxAnzahlPruefer,99);


    Refresh(Einstellungen);
    ClearCollect(colSystemSettings,
        Filter(Einstellungen,
            Load=true
        )
    );
    Set(SystemSettingsLoaded,true);
    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
//---------------------------------------------
If( SystemSettingsLoaded &&
    !MySettingsLoaded,
    Set (_ShowLoadingHint,true);
    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="Einstellungen").Wert
    );

    Set(MyApp,
        First(Filter(PowerAppsforMakers.GetApps().value,
        properties.displayName=MyAppName))
    );

    Set(MyAppID,MyApp.name);

    Set(MyAppUsedDays,
        DateValue(Now())-
        DateValue(MyApp.properties.createdTime)
    );

    // App Version vom SharePoint laden
    UpdateContext({tmpString:MyAppName & ".Version"});
    UpdateContext({MyVersionFromSystemSettings:
        LookUp(colSystemSettings,Title=tmpString,Wert)
    });
    // wenn die lokale Version nicht mit der Version vom 
    // SharePoint übereinstimmt, 
    // dann soll ein Hinweis angezeigt werden
    UpdateContext({_ShowVersionHint:false});
    If(MyVersion<>MyVersionFromSystemSettings, 
        UpdateContext({_ShowVersionHint:true})
    ); 

    // Environment 
    // Die angegebene Datei ist nicht wichtig!
    // Der Flow macht nichts,
    // liefert nur den Wert einer Umgebungsvariablen.
    // Das kann PowerApps zwar auch, wird aber als
    // Premium-Dienst gewertet und ist daher teuer!
    
    Set(_Environment, 
        'pobvolService:GetTeamv2'.Run().response
    );

    //Team
    Set(MyTeam,
        Last(ForAll(TrimEnds(ForAll(Split(_Environment,"/"),
        {Result: ThisRecord.Value})), 
        {Result: ThisRecord.Value})).Result
    );

    Set(DokumenteRootFolder, 
        'pobvolService:GetDocumentRootFolder'.
        Run().
        response
    );
    Set(DokumenteRootFolder, 
        Last(
            Split(DokumenteRootFolder,"/")
        ).Value
    );

    /*
    Set(MyGroups,
        'Office365-Gruppen'.ListGroups().value
    );
    */

    Set(MyTeamId,
        First(
            Filter(
                'Office365-Gruppen'.ListGroups().value,
                mailNickname=MyTeam
            )
        ).id
    );

    //Domäne
    Set(MyDomain,
        First(ForAll(Split(
            Lower(TrimEnds(Right(_Environment,
            Len(_Environment)-8)))
            ,".sharepoint"), {Result: ThisRecord.Value})
        ).Result
    );

    //User Domäne ermitteln aus E-Mail
    /*
    Set(_userDomain,
        Lower(TrimEnds(Right(User().Email,
        Len(User().Email) - Find("@", User().Email))))
    );
    If(!IsBlank(MyDomain),
        Set(_userDomain,MyDomain)
    );
    */

    //Set(_dateSelected,Today());

    /*
    Set(_firstDayOfMonth, 
        DateAdd(_dateSelected, 
        1 - Day(_dateSelected), TimeUnit.Days)
    );
    */

    /*
    Set(_firstDayInView, 
        DateAdd(_firstDayOfMonth, 
        -(Weekday(_firstDayOfMonth - 1) - 2 + 1), 
        TimeUnit.Days)
    );
    */
    
    /*
    Set(_lastDayOfMonth, 
        DateAdd(
            DateAdd(
                _firstDayOfMonth, 
                1, 
                TimeUnit.Months), 
            -1, 
            TimeUnit.Days
        )
    );
    */

    //Set(_minDate,_firstDayOfMonth);
    //Set(_maxDate,_lastDayOfMonth);

    /*
    UpdateContext({MyDomainTesterSPValue:
        LookUp(
            colSystemSettings;
            Title="pssService License User";
            Wert)
    });; 
    */

    Set(pssFakturaAddOnActive,false);
    If(Lower(LookUp(colSystemSettings,
            Title="pssFaktura AddOn.Active",
            Wert))="true",
        Set(pssFakturaAddOnActive,true)
    );

    Set(Currency,
        LookUp(colSystemSettings,Title="Currency",Wert)
    );

    Set(MeetingOrganizer,
        LookUp(colSystemSettings,
        Title="Meeting Organizer (E-Mail)",Wert)
    );

    UpdateContext(
        {
            _Title:User().Email
            &".Termine verschicken aus Kalender"
        }
    );
    Set(SharedCalendar,
        LookUp(EinstellungenBenutzer,
        Title=_Title, Wert)
    );

    UpdateContext(
        {
            _Title:User().Email
            &".Termine verschicken aus Kalender (Name)"
        }
    );
    Set(SharedCalendarName,
        LookUp(EinstellungenBenutzer,
        Title=_Title,Wert)
    );

    Set(TeamEmail,
        LookUp(colSystemSettings,
        Title="Service orders.Shared mailbox",Wert)
    );

    Set(Arbeitstaganfang,
        LookUp(colSystemSettings,
        Title="Arbeitsaganfang um x Uhr", Wert)
    );
    Set(ArbeitstaganfangUhrzeit, 
        DateAdd(DateTimeValue("01.01.1900 00:00", "DE"),
        Value(Arbeitstaganfang),TimeUnit.Hours)
    );
    Set(Arbeitstagende,
        LookUp(colSystemSettings,
        Title="Arbeitstagende um y Uhr", Wert)
    );
    Set(ArbeitstagendeUhrzeit, 
        DateAdd(DateTimeValue("01.01.1900 00:00", "DE"),
        Value(Arbeitstagende),TimeUnit.Hours)
    );
    Set(StandardTerminDauer,
        LookUp(colSystemSettings,
        Title="Standard Termindauer in Minuten", Wert)
    );

    //Notify("Lade den Termintext aus SP-Liste Einstellungen");
    Set(UVVTerminText,
        LookUp(colSystemSettings,
        Title="Service orders.Description.Text", Wert)
    );

    Set(UVVTerminanKundensenden,
        LookUp(colSystemSettings,
        Title="Service reports.Send reports to customers", Wert)
    );

    Set(WeblinkUVVPruefung,
        LookUp(colSystemSettings,
        Title="pssService Phone.Weblink", Wert)
    );
    
    Set(txtBingMapsKey,
        LookUp(colSystemSettings,
        Title="txtBingMapsKey", Wert)
    );

    Set(TeamsChannelFolder,
        LookUp(colSystemSettings,
        Title="SharePoint.TeamsChannelFolder", Wert)
    );

    Set(VertraegeFolder,
        LookUp(colSystemSettings,
        Title="Contracts are saved in folder", Wert)
    );

    Set(ServiceberichteFolder,
        LookUp(colSystemSettings,
        Title="Service reports are saved in folder", Wert)
    );

    UpdateContext(
        {_Title:User().Email&".ArbeitstaganfangUserStunde"}
    );
    Set(ArbeitstaganfangUserStunde,
        LookUp(EinstellungenBenutzer,Title=_Title).Wert);
    
    UpdateContext(
        {_Title:User().Email&".ArbeitstaganfangUserMinuten"}
    );
    Set(ArbeitstaganfangUserMinuten,
        LookUp(EinstellungenBenutzer,Title=_Title).Wert);
    
    UpdateContext(
        {_Title:User().Email&".StartingPoint"}
    );
    Set(StartingPointUser,
        LookUp(EinstellungenBenutzer,Title=_Title).Wert);

    UpdateContext({_Title:
        User().Email&"."&MyAppName&".MyFontSize"
    });
    Set(MyFontSize,
        Value(LookUp(EinstellungenBenutzer,Title=_Title).Wert)
    );
    If(IsBlank(MyFontSize)||MyFontSize<10,
        Set(MyFontSize,10)
    );

    //Notify("Lade den Power BI Report - Link aus SP-Liste Einstellungen");
    Set(PowerBIReport,
        LookUp(colSystemSettings,
        Title="PowerBIReport", Wert)
    );

    Set(MySettingsLoaded, true);
    Set (_ShowLoadingHint,false);
    Set(FooterText,"");

);
//------------------------------------------------
// Prüfer vom SharePoint laden
If( (!PrueferLoaded||IsBlank(PrueferLoaded)) 
    && !IsBlank(MyTeamId),
    Set (_ShowLoadingHint,true);
    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="die Benutzer").Wert
    );
    Set(PrueferLoaded,true);

    //Alle Tenant Benutzer
    ClearCollect(colUser,
        Office365Users.SearchUser()
    );

    //Nur die Mitglieder der Gruppe
    //Set(Groups,
    //    'Office365-Gruppen'.ListGroups()
    //);
    Set(GroupMembers,
        'Office365-Gruppen'.ListGroupMembers(MyTeamId)
    );
    ClearCollect(colGroupMembers,
        AddColumns(GroupMembers.value As aSource,
            AccountEnabled,
            LookUp(colUser,Id=aSource.id,AccountEnabled)
        )
    );

    //Prüfer / Benutzer 
    UpdateContext({_value1:
    MyDomain & "." & MyTeam & ".Tester"});
    ClearCollect(colPruefer,
        AddColumns(
            ShowColumns(
                Filter(colSystemSettings,
                    Title=_value1
                ),
                ID,
                Wert
            )
            ,
            Pruefername, 
            LookUp(colUser,Mail=Wert,DisplayName),
            PrueferMail, Wert,
            UserPrincipalName,
            LookUp(colUser,Mail=Wert,UserPrincipalName)
        )
    );
    
    Set(AnzahlPruefer,
        CountRows(Filter(colPruefer,!IsBlank(Wert)))
    ); 

    Set(UserEmail,
        If("#" in User().Email,
            Last(ForAll(Split(User().Email,"#"), {Result: ThisRecord.Value})).Result
        ,//Else
            User().Email
        )
    );
    If(UserEmail in Concat(colPruefer,Wert),
        Set(IsPruefer,true)
    ,//Else
        Set(IsPruefer,false)
    );

    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
//---------------------------------------------
// Admins vom SharePoint laden
If( (!AdminsLoaded||IsBlank(AdminsLoaded))
    && !IsBlank(MyTeamId),
    Set (_ShowLoadingHint,true);
    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="die Admins").Wert
    );
    Set(AdminsLoaded,true);

    UpdateContext({_value1:
        MyDomain & "." & MyTeam & ".Admin"});

    ClearCollect(colAdmins,
        AddColumns(
            ShowColumns(
                Filter(colSystemSettings,
                    Title=_value1
                ),
                ID,
                Wert
            )
            ,
            Benutzername, "",
            //LookUp(colUser;Mail=Wert;DisplayName);
            Benutzermail, Wert
        )
    );

    Set(UserEmail,
        If("#" in User().Email,
            Last(ForAll(Split(User().Email,"#"), {Result: ThisRecord.Value})).Result
        ,//Else
            User().Email
        )
    );
    If(UserEmail in Concat(colAdmins,Wert),
        Set(IsAdmin,true)
    ,//Else
        Set(IsAdmin,false)
    );

    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
//---------------------------------------------
//Darf Serviceauftraege bearbeiten
If( (!EditorsServiceorderLoaded||
    IsBlank(EditorsServiceorderLoaded))
    && !IsBlank(MyTeamId),
    Set (_ShowLoadingHint,true);
    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="Wer erstellt und bearbeitet Serviceaufträge")
        .Wert
    );
    Set(EditorsServiceorderLoaded,true);

    UpdateContext({_value1:
        MyDomain & "." & MyTeam & ".EditorServiceorder"});

    ClearCollect(colEditorsServiceorders,
        AddColumns(
            ShowColumns(
                Filter(colSystemSettings,
                    Title=_value1
                ),
                ID,
                Wert
            )
            ,
            Benutzername, "",
            //LookUp(colUser;Mail=Wert;DisplayName);
            Benutzermail, Wert
        )
    );

    Set(UserEmail,
        If("#" in User().Email,
            Last(ForAll(Split(User().Email,"#"), {Result: ThisRecord.Value})).Result
        ,//Else
            User().Email
        )
    );
    If(User().Email in Concat(colEditorsServiceorders,Wert),
        Set(IsEditorServiceorders,true),
        Set(IsEditorServiceorders,false)
    );
    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
//---------------------------------------------
//Versionshinweis
/*
If(_ShowVersionHint&&OnlineStatus=1;
    Set(_ShowVersionHint;false);;
    //Navigate(AppVersion)
);;
*/
//------------------------------------------------
//Lizenzhinweis anzeigen
/*
If(ShowKeineLizenz&&OnlineStatus=1;
    Set(ShowKeineLizenz,false);
    //Navigate(AppLizenz)
);;
*/
//------------------------------------------------
//Set(_ShowAppStart,false);