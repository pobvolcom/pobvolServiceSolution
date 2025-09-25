//pssServiceBoard ScreenPruefer ScreenPrueferREAD.fx
//Power Apps Language: en-US
//------------------------------------------------
//-----------------------------------------------------
//User und Prüfer ermitteln
If(IsBlank(ReadPruefer)||ReadPruefer||!PrueferLoaded,
    
    Set (_ShowLoadingHint,true);
    //Notify("Lade Benutzer...");

    //-----------------------------------------------------
    If(IsBlank(PrueferSortDescending),
        UpdateContext({PrueferSortDescending:false});
    );
    //-----------------------------------------------------
    Set(ReadPruefer,false);
    Set(PrueferLoaded,true);
    //-----------------------------------------------------
    //Auswahl
    ClearCollect(colUserMenu,
        {
            Title:"die Benutzer",
            TitleText:
            LookUp(MyUVVSettings,Title="Prüfer (Mehrzahl)").Wert,
            Tab:1
        },
        {
            Title:"die Admins",
            TitleText:
            LookUp(MyUVVSettings,Title="die Admins").Wert,
            Tab:2
        },
        {
            Title:"Serviceauftraege",
            TitleText:
            LookUp(MyUVVSettings,Title="Wer erstellt und bearbeitet Serviceaufträge").Wert,
            Tab:3
        }
    );
    //-----------------------------------------------------
    //Alle Tenant Benutzer
    ClearCollect(colUser,
        Filter(Office365Users.SearchUser(),AccountEnabled)
    );
    //-----------------------------------------------------
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
    //-----------------------------------------------------
    //Prüfer / Benutzer 
    UpdateContext({_value1:
    MyDomain & "." & MyTeam & ".Tester"});
    ClearCollect(colPruefer,
        AddColumns(
            ShowColumns(
                Filter(Einstellungen,
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
    //-----------------------------------------------------
    //Admins
    UpdateContext({_value1:
    MyDomain & "." & MyTeam & ".Admin"});
    ClearCollect(colAdmins,
        AddColumns(
            ShowColumns(
                Filter(Einstellungen,
                    Title=_value1
                ),
                ID,
                Wert
            )
            ,
            Benutzername, 
            LookUp(colUser,Mail=Wert,DisplayName),
            Benutzermail, Wert
        )
    );
    //-----------------------------------------------------
    //Darf Serviceauftraege bearbeiten
    UpdateContext({_value1:
    MyDomain & "." & MyTeam & ".EditorServiceorder"});
    ClearCollect(colEditorsServiceorders,
        AddColumns(
            ShowColumns(
                Filter(Einstellungen,
                    Title=_value1
                ),
                ID,
                Wert
            )
            ,
            Benutzername, 
            LookUp(colUser,Mail=Wert,DisplayName),
            Benutzermail, Wert
        )
    );
    //-----------------------------------------------------
    Set(AnzahlPruefer,
        CountRows(Filter(colPruefer,!IsBlank(Wert)))
    ); 
    //-----------------------------------------------------
    If(User().Email in Concat(colPruefer,Wert),
        Set(IsPruefer,true),
        Set(IsPruefer,false)
    );
    //-----------------------------------------------------
    If(User().Email in Concat(colAdmins,Wert),
        Set(IsAdmin,true),
        Set(IsAdmin,false)
    );
    //-----------------------------------------------------
    If(User().Email in Concat(colEditorsServiceorders,Wert),
        Set(IsEditorServiceorders,true),
        Set(IsEditorServiceorders,false)
    );
    //-----------------------------------------------------
    Set (_ShowLoadingHint,false);
    //-----------------------------------------------------

);
