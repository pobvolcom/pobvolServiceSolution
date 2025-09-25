//pssServicePhone Startscreen TimerReadAuftraege1
//Power Apps Language: en-US
//------------------------------------------------
If(Connection.Connected,
    Set (OnlineStatus, 1)
,// Else    
    Set (OnlineStatus, 0)
);
//---------------------------------------------
If(IsBlank(ReadVon),
    Set(Tag,Today());
    Set(ReadVon,
        DateAdd(
            Today(),
            -(Weekday(Today(),StartOfWeek.MondayZero)),
            TimeUnit.Days
        )
    );
    Set(ReadBis,DateAdd(ReadVon,6,TimeUnit.Days));
);
//---------------------------------------------
If(IsBlank(ReadSixMonthsVon),
    Set(ReadSixMonthsVon,
        //Today()-180
        EDate(Today(),-5)	
    );
    If(Day(ReadSixMonthsVon)>1,
        Set(ReadSixMonthsVon,
            Date(
                Year(ReadSixMonthsVon),
                Month(ReadSixMonthsVon),
                1
            )
        )
    );
    Set(ReadSixMonthsBis,
        //Today()
        EOMonth(Today(),0)
    );
);
//-----------------------------------------------------
If(OnlineStatus=1 && 
    LanguageLoaded &&
    (IsBlank(ReadAuftraege)||ReadAuftraege &&
    MyItemsDefaultLoaded),

    //-----------------------------------------------------
    If(IsBlank(AuftraegeSortDescending),
        UpdateContext({AuftraegeSortDescending:false});
    );
    //-----------------------------------------------------
    //offene Aufträge des festgelegten Zeitraums
    Set(_ShowLoadingHint,true);

    Set(FooterText,"Loading service orders and reminders from SP");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});

    Refresh(Serviceauftraege);
    //-----------------------------------------------------
    //Alle Aufträge und Erinnerungen des Zeitraums laden
    ClearCollect(tmpAuftraege0,
        Filter(Serviceauftraege,
            Anfangszeit >= ReadVon &&
            Anfangszeit <= ReadBis
        )
    );
    //-----------------------------------------------------
    //Status der Aufträge ermitteln
    //StatusKey ist nicht leer, wenn es mindestens eine Position ohne Servicevorgang gibt
    //Heisst: Wenn StatusKey = Blank(), dann ist der Auftrag abgeschlossen
    ClearCollect(tmpAuftraege1,
        AddColumns(
            tmpAuftraege0 As aSource,
            StatusKey,
                LookUp(ServiceauftraegeP,
                    ParentID=aSource.ID && 
                    PruefdatenGefunden=Blank(),
                    ID
                )
        )
    );
    Clear(tmpAuftraege0);
    //-----------------------------------------------------
    //Nur mit offenen Aufträgen weitermachen
    ClearCollect(colAuftraege1,
        Filter(tmpAuftraege1,
            !IsBlank(StatusKey)
        )
    );
    Clear(tmpAuftraege1);
    //-----------------------------------------------------
    //Zugehörige Positionen laden
    Clear(MyAuftraegeP);
    ForAll(colAuftraege1 As aSource,
        Collect(MyAuftraegeP,
            ShowColumns(
                Filter(ServiceauftraegeP,ParentID=aSource.ID),
                ID,KEY,ParentID,Pos,KDNR,INVNR,PruefdatenGefunden
            )
        )
    );
    //-----------------------------------------------------
    //Zugehörige Kunden laden
    ClearCollect(MyAuftraegeKDNR,
        GroupBy(colAuftraege1,KDNR,ID)
    );
    Clear(MyAuftraegeKunden);
    ForAll(MyAuftraegeKDNR As aSource,
        Collect(MyAuftraegeKunden,
            ShowColumns(
                Filter(Servicekunden,
                    KDNR=aSource.KDNR
                ),
                ID,KDNR,Kunde,Strasse,Plz,Kundenort,Land,GPSLocation,Bemerkungen,Ansprechpartner,Telefon,'E-Mail',Sprache,Kommunikationsart,GPSLocationBreitengrad,GPSLocationLaengengrad
            )
        )
    );
    //-----------------------------------------------------
    //Die zugeordneten Geräte laden
    ClearCollect(MyAuftraegeKDNRINVNR,
        GroupBy(MyAuftraegeP,KDNR,INVNR,ID)
    );
    Clear(MyAuftraegeInventar);
    ForAll(MyAuftraegeKDNRINVNR As aSource,
        Collect(MyAuftraegeInventar,
            Filter(Kundeninventar,
                KDNR=aSource.KDNR,
                INVNR=aSource.INVNR
            )
        )
    );
    //-----------------------------------------------------
    ClearCollect(MyAuftraege,
        SortByColumns(
            AddColumns(
                
                colAuftraege1 As aSource,

                ServiceartText,
                    LookUp(
                    colFormular,
                    Serviceart=aSource.Serviceart,
                    ServiceartText),

                ChecklisteText,
                    LookUp(
                    colFormular,
                    Serviceart=aSource.Serviceart &&
                    Geraetetyp=aSource.GeraetetypCode &&
                    Checkliste=aSource.Inventarart,
                    ChecklisteText
                    ),

                GeraetetypText,
                    LookUp(colFormular,
                    Serviceart=aSource.Serviceart &&
                    Geraetetyp=aSource.GeraetetypCode,
                    GeraetetypText),

                Tag,Day(aSource.Anfangszeit),
                Monat,Month(aSource.Anfangszeit),
                Jahr,Year(aSource.Anfangszeit),
                Pruefzeitpunkt,
                    Month(aSource.Anfangszeit)
                    &"/"
                    &Right(Year(aSource.Anfangszeit),2),

                Dauer,
                    24/1*(aSource.Endzeit-aSource.Anfangszeit)*60,

                Kunde,
                    LookUp(
                        MyAuftraegeKunden,
                        KDNR=aSource.KDNR,
                        Kunde
                    ),
                Kundenort,
                    LookUp(
                        MyAuftraegeKunden,
                        KDNR=aSource.KDNR,
                        Kundenort
                    ),

                AuftragDevices,
                    Value(aSource.HISTORIEDE),

                IncludedDevices,
                    aSource.HISTORIEEN

            ),
            "Anfangszeit",  
                If(AuftraegeSortDescending,
                SortOrder.Descending,SortOrder.Ascending),
            "Kunde", SortOrder.Ascending
        )
    );
    Clear(colAuftraege1);
    //-----------------------------------------------------
    //Nur was in der aktuellen Woche gemacht werden muss, wird lokal gespeichert
    //Wenn sich der Techniker Dinge aus der Vergangenheit oder Zukunft anschaut,
    //dann wird das zwar geladen, aber nicht gespeichert!
    
    If(MyDeviceType="Mobile" && Today()>=ReadVon && Today()<=ReadBis,

        Set(FooterText,"Saving MyAuftraege to cache");
        Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
        SaveData (MyAuftraege,"MyAuftraege");    

        Set(FooterText,"Saving MyAuftraegeP to cache");
        Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
        SaveData (MyAuftraegeP,"MyAuftraegeP");    

        Set(FooterText,"Saving MyAuftraegeKunden to cache");
        Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
        SaveData (MyAuftraegeKunden,"MyAuftraegeKunden");    

        Set(FooterText,"Saving MyAuftraegeInventar to cache");
        Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
        SaveData (MyAuftraegeInventar,"MyAuftraegeInventar");
    );
    //-----------------------------------------------------
    If(CountRows(MyAuftraege)>0,
        If(IsBlank(AuftragSelected)||
           !AuftragSelected||
           IsBlank(SAuftragId),

            Set(SAuftragId,First(MyAuftraege).ID);
            Set(AuftragSelected,true)

        ,//Else
        
            Set(AuftragSelected,true)
        )
    ,//Else
        Set(SAuftragId,Blank());
        Set(AuftragSelected,false)
    );
    //-----------------------------------------------------
    Set(_ShowLoadingHint,false);
    Set(FooterText,"");
    Set(ReadAuftraege,false);
    Set(ReadMyWeek,true);
    Set(ReadAuftraegeChart,true);
    //-----------------------------------------------------
);
//-----------------------------------------------------
//Serviceaufträge vom Cache laden, 
//wenn Offline und Mobile und noch nicht geladen
If( OnlineStatus=0 && 
    LanguageLoaded &&
    (IsBlank(ReadAuftraege)||ReadAuftraege) &&
    MyDeviceType="Mobile" &&
    MyItemsDefaultLoaded,

    //-----------------------------------------------------
    If(IsBlank(AuftraegeSortDescending),
        UpdateContext({AuftraegeSortDescending:false});
    );
    //-----------------------------------------------------
    Set(_ShowLoadingHint,true); 

    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="Lade Termine vom Cache").Wert
    );

    Set(FooterText,"Loading MyAuftraege from cache");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
    Clear(MyAuftraege);
    LoadData(MyAuftraege,"MyAuftraege",true);
    
    Set(FooterText,"Loading MyAuftraegeP from cache");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
    Clear(MyAuftraegeP);
    LoadData(MyAuftraegeP,"MyAuftraegeP",true);

    Set(FooterText,"Loading MyAuftraegeKunden from cache");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
    Clear(MyAuftraegeKunden);
    LoadData(MyAuftraegeKunden,"MyAuftraegeKunden",true);
    
    Set(FooterText,"Loading MyAuftraegeInventar from cache");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
    Clear(MyAuftraegeInventar);
    LoadData(MyAuftraegeInventar,"MyAuftraegeInventar",true);
    //-----------------------------------------------------
    If(CountRows(MyAuftraege)>0,
        If(IsBlank(AuftragSelected)||
           !AuftragSelected||
           IsBlank(SAuftragId),
            Set(SAuftragId,First(MyAuftraege).ID);
            Set(AuftragSelected,true)
        ,//Else
            Set(AuftragSelected,true)
        )
    ,//Else
        Set(SAuftragId,Blank());
        Set(AuftragSelected,false)
    );
    //-----------------------------------------------------
    Set(_ShowLoadingHint,false);
    Set(FooterText,"");
    Set(ReadAuftraege,false);
    Set(ReadMyWeek,true);
    Set(ReadAuftraegeChart,true);
    //-----------------------------------------------------
);
//-----------------------------------------------------
//MyWeek
If(ReadMyWeek && LanguageLoaded,
    
    Set(_ShowLoadingHint,true);
    Set(FooterText,"MyWeek");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});

    If(IsBlank(MyWeekSortDescending),
        Set(MyWeekSortDescending,true)
    );

    Clear(tmpWeek);

    //------------------------------------------------
    //Serviceaufträge und Erinnerungen hinzufügen
    ForAll(

        ShowColumns(
            Filter(MyAuftraege,
                Anfangszeit >= ReadVon && 
                Anfangszeit <= ReadBis
            ),
            ID,KEY,Anfangszeit,Pruefer,PrueferName
        ) As aSource,

        If(User().FullName in aSource.PrueferName,
            Collect(tmpWeek,
            {
                ID:Value(Text(CountRows(tmpWeek)+1)), 
                KEY:aSource.KEY,
                Kennung:"A",
                Datum:DateValue(aSource.Anfangszeit)
            })
        )
    );
    //------------------------------------------------
    //Servicevorgänge des Users für den ausgewählten Zeitraum vom SP laden
    If(OnlineStatus=1,
    ForAll(
        RenameColumns(
            ShowColumns(
                Filter(Servicevorgaenge,
                    Pruefer=User().FullName 
                    && Pruefdatum >= ReadVon 
                    && Pruefdatum <= ReadBis 
                ),
                ID,KEY,Pruefdatum,Pruefer
            ),
            field_4,Pruefdatum
        ) As aSource,
        Collect(tmpWeek,
        {
            ID:Value(Text(CountRows(tmpWeek)+1)), 
            KEY:aSource.KEY,
            Kennung:"S",
            Datum:DateValue(aSource.Pruefdatum)
        })
    ));
    //------------------------------------------------
    //Servicevorgänge des Users für den ausgewählten Zeitraum vom Cache laden
    ForAll(
        ShowColumns(
            Filter(MyItems,
                Pruefer=User().FullName,
                //DateValue(Text(Pruefdatum;"yyyy.mm.dd";"de"))
                DateValue(Text(Pruefdatum,"yyyy.mm.dd","de"))>=DateValue(ReadVon),
                DateValue(Text(Pruefdatum,"yyyy.mm.dd","de"))<=DateValue(ReadBis)
                //Pruefdatum>=Text(ReadVon;"yyyymmdd";"de")
                //Text(Pruefdatum;"yyyymmdd";"de")<=Text(ReadBis;"yyyymmdd";"de")
            ),
            ID,KEY,Pruefdatum,Pruefer
        ) As aSource,
        //Wenn der Vorgang noch nicht vorhanden ist, ...
        If(CountIf(tmpWeek,Kennung="S"&&KEY=aSource.KEY)=0,
        Collect(tmpWeek,
        {
            ID:Value(Text(CountRows(tmpWeek)+1)), 
            KEY:aSource.KEY,
            Kennung:"S",
            Datum:DateValue(Text(aSource.Pruefdatum,"yyyy.mm.dd","de"))
        }))
    );
    //------------------------------------------------
    //Rückreisen des Users für den ausgewählten Zeitraum vom SP laden
    If(OnlineStatus=1,
    ForAll(
        ShowColumns(
            Filter(Fahrtbericht,
                Title="Rückreise",
                Pruefer=User().FullName,
                Reisedatum >= ReadVon && 
                Reisedatum <= ReadBis 
            ),
            ID,KEY,Reisedatum,Pruefer
        ) As aSource,
        Collect(tmpWeek,
        {
            ID:Value(Text(CountRows(tmpWeek)+1)), 
            KEY:aSource.KEY,
            Kennung:"R",
            Datum:DateValue(aSource.Reisedatum)
        })
    ));
    //------------------------------------------------
    //Rückreisen des Users für den ausgewählten Zeitraum vom Cache laden
    ForAll(
        ShowColumns(
            Filter(MyItemsFahrtbericht,
                Titel="Rückreise",
                Pruefer=User().FullName,
                //Reisedatum >= ReadVon && 
                //Reisedatum <= ReadBis 
                DateValue(Text(Reisedatum,"yyyy.mm.dd","de"))>=DateValue(ReadVon),
                DateValue(Text(Reisedatum,"yyyy.mm.dd","de"))<=DateValue(ReadBis)
            ),
            KEY,Reisedatum,Pruefer
        ) As aSource,
        //Wenn die Rückreise noch nicht vorhanden ist, ...
        If(CountIf(tmpWeek,Kennung="R"&&KEY=aSource.KEY)=0,
        Collect(tmpWeek,
        {
            ID:Value(Text(CountRows(tmpWeek)+1)), 
            KEY:aSource.KEY,
            Kennung:"R",
            //Datum:DateValue(aSource.Reisedatum)
            Datum:DateValue(Text(aSource.Reisedatum,"yyyy.mm.dd","de"))
        }))
    );
    //------------------------------------------------
    //Jetzt gruppieren nach Datum
    ClearCollect(MyWeek,
        SortByColumns(
            AddColumns(
                GroupBy(tmpWeek,Datum,REST),
                Auftraege,CountIf(REST,Kennung="A"),
                Vorgaenge,CountIf(REST,Kennung="S"),
                Rueckreisen,CountIf(REST,Kennung="R")
            ),
            "Datum",
                If(MyWeekSortDescending,
                    SortOrder.Descending
                ,//Else
                    SortOrder.Ascending
                )
        )
    );
    Clear(tmpWeek);
    //-----------------------------------------------------
    //Wenn was gefunden wurde, den ersten Eintrag auswählen
    //Wenn was für den aktuellen Tag gefunden wurde, den aktuellen Tag auswählen

    If(CountRows(MyWeek)>0,

	    Set(Tag,First(MyWeek).Datum);

        If(!IsBlank(LookUp(MyWeek,Datum=Today())),
            Set(Tag,Today())
        )

    ,//Else

        Set(Tag,Blank())

    );
    //-----------------------------------------------------
    Set(_ShowLoadingHint,false);
    Set(ReadMyWeek,false);
    Set(FooterText,"");    
    
);
//-----------------------------------------------------
//Auftragschart bauen
If( LanguageLoaded &&
    (IsBlank(ReadAuftraegeChart)||ReadAuftraegeChart),

    //offene Aufträge des festgelegten Zeitraums
    Set(_ShowLoadingHint,true);
    Set(FooterText,"Creating collection MyAuftraegeChart");
    Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
    ClearCollect(MyAuftraegeChart,
        SortByColumns(
            AddColumns(
                GroupBy(
                    /*Filter(MyAuftraege;
                        Anfangszeit >= ReadVon &&
                        Anfangszeit <= ReadBis
                    );*/
                    MyAuftraege,
                    Jahr,
                    Monat,
                    Tag,
                    ByDummy
                ),
                Pruefzeitpunkt, Monat&"/"&Right(Jahr,2),
                '# per Day',CountRows(ByDummy),
                TagShort,
                    If(MyLanguage="de",
                        Text(Tag&"."&Monat)
                    ,//Else
                        Text(Monat&"/"&Tag)
                    )
            ),
            "Jahr",SortOrder.Ascending,
            "Monat",SortOrder.Ascending,
            "Tag",SortOrder.Ascending
        )
    );
    Set(_ShowLoadingHint,false);
    Set(FooterText,"");
    Set(ReadAuftraegeChart,false);
);
//---------------------------------------------
//Distanz und Dauer laden
If( LanguageLoaded && 
    (IsBlank(LoadDistance)||LoadDistance),

    Set(_ShowLoadingHint,true);

    Set(LoadDistance,false);

    If(OnlineStatus=1,
        Set(FooterText,"Loading distance from SP");
        Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
        ClearCollect(colDistanceGrouped,
            GroupBy(
                ShowColumns(
                    Filter(Fahrtbericht,
                        Pruefer=User().FullName
                    ),
                    Abfahrtsort,
                    Zielort,
                    Reisezeit,
                    Distanz
                ),
                Abfahrtsort,Zielort,REST
            )
        );
        ClearCollect(tmpDistance,
            AddColumns(colDistanceGrouped,
                Distance,
                    //Max(Concat(REST;Distanz;";"));
                    Last(REST).Distanz,
                Duration,
                    //Max(Concat(REST;Reisezeit;";"))
                    Last(REST).Reisezeit
            )
        );
        ClearCollect(colDistance,
            ShowColumns(
                tmpDistance,
                Abfahrtsort,
                Zielort,
                Distance,
                Duration
            )
        );
        //Clear(colDistanceGrouped);;
        //Clear(tmpDistance);;
        //Daten lokal speichern wenn Mobile
        If(MyDeviceType="Mobile",
            Set(FooterText,"Saving distance to cache");
            Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
            SaveData(colDistance,"Distance")
        )
    ,//Else
        If(MyDeviceType="Mobile",
            Set(FooterText,"Loading distance from cache");
            Collect(colLog,{Text:Text(Now(),"dd.mm.yyyy hh:mm:ss","de")&" "&FooterText});
            Clear(colDistance);
            LoadData(colDistance,"Distance",true);
        )
    );
    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
//------------------------------------------------
If(LanguageLoaded &&
    (IsBlank(ReadNewReports)||ReadNewReports) && 
    OnlineStatus=1,
    
    Set (_ShowLoadingHint,true);
    //------------------------------------------------
    Set(FooterText,"Checking for new reports");
    //------------------------------------------------
    Set(ReadNewReports,false);
    //------------------------------------------------
    ClearCollect(colBerichtedesPruefers,
        AddColumns(
            GroupBy(
                ShowColumns(
                    Filter(Serviceberichte,
                        Pruefer=User().FullName
                    ),
                    Pruefer,ID
                ),
                Pruefer,
                Rest
            ),
            Anzahl,CountRows(Rest)
        )
    );
    Set(CountofUVVBerichte,
        LookUp(colBerichtedesPruefers,
            Pruefer=User().FullName,
            Anzahl
        )
    );
    //-----------------------------------------------------
    ClearCollect(colNichtGenehmigteBerichte,
        AddColumns(
            GroupBy(
                Filter(Serviceberichte,Pruefer=User().FullName),
                Genehmigt,
                Rest
            ),
            Anzahl,CountRows(Rest),
            Prozent,
            If(MyLanguage="de",
                Text(CountRows(Rest)/CountofUVVBerichte*100,
                "[$-de-DE]0%")
            ,//Else
                Text(CountRows(Rest)/CountofUVVBerichte*100,
                "[$-en-US]0%")
            )
        )
    );
    //-----------------------------------------------------
    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
//------------------------------------------------
If(LanguageLoaded &&
    (IsBlank(ReadBerichte)||ReadBerichte) && 
    OnlineStatus=1,
    
    Set (_ShowLoadingHint,true);
    //------------------------------------------------
    Set(FooterText,
        LookUp(MyUVVSettings,
        Title="Serviceberichte").Wert
    );
    //------------------------------------------------
    Set(ReadBerichte,false);
    //------------------------------------------------
    //Berichte pro Monat
    //10 Berichte pro Tag sind 200 Berichte pro Monat
    //Da könnte man denken, dass das schnell knapp wird mit der maximalen Anzahl von 500 Datensätzen vom SP. Aber hier wird gruppiert eingelesen!
    ClearCollect(
        coltmp1,
        SortByColumns(
            AddColumns(
                GroupBy(
                    ShowColumns(
                        Filter(
                            ArchivServiceberichte,
                            Pruefer=User().FullName,
                            Pruefdatum>=ReadSixMonthsVon&&
                            Pruefdatum<=ReadSixMonthsBis
                        ),
                        Pruefjahr,
                        Pruefmonat,
                        Pruefer,
                        ID
                    ),
                    Pruefjahr,
                    Pruefmonat,
                    Pruefer,
                    ByPruefmonat
                ),
                Pruefpunkt,
                    Text(Date(Pruefjahr,Pruefmonat,1),"mmm"),
                BerichteproMonat,
                    CountRows(ByPruefmonat)
            ),
            "Pruefjahr",SortOrder.Ascending,
            "Pruefmonat",SortOrder.Ascending
        )
    );
    ClearCollect(
        coltmp2,
        SortByColumns(
            AddColumns(
                GroupBy(
                    ShowColumns(
                        Filter(
                            Serviceberichte,
                            Pruefer=User().FullName,
                            Pruefdatum>=ReadSixMonthsVon&&
                            Pruefdatum<=ReadSixMonthsBis
                        ),
                        Pruefjahr,
                        Pruefmonat,
                        Pruefer,
                        ID
                    ),
                    Pruefjahr,
                    Pruefmonat,
                    Pruefer,
                    ByPruefmonat
                ),
                Pruefpunkt,
                    Text(Date(Pruefjahr,Pruefmonat,1),"mmm"),
                BerichteproMonat,
                    CountRows(ByPruefmonat)
            ),
            "Pruefjahr",SortOrder.Ascending,
            "Pruefmonat",SortOrder.Ascending
        )
    );
    ClearCollect(coltmp3,coltmp1,coltmp2);
    Clear(coltmp1);
    Clear(coltmp2);
    //-----------------------------------------------------
    ClearCollect(
        colBerichteProMonat,
        SortByColumns(
            AddColumns(
                GroupBy(
                    coltmp3,
                    Pruefpunkt,
                    Pruefjahr,
                    Pruefmonat,
                    ByPruefpunkt
                ),
                'Berichte pro Monat',
                    Sum(ByPruefpunkt,BerichteproMonat),
                Ausgabe,
                    If(Sum(ByPruefpunkt,BerichteproMonat)<10,
                        10
                    ,//Else
                        Sum(ByPruefpunkt,BerichteproMonat)
                    )

            ),
            "Pruefjahr",SortOrder.Ascending,
            "Pruefmonat",SortOrder.Ascending
        )
    );
    //-----------------------------------------------------
    /*
    ClearCollect(colBerichteProMonat;
        SortByColumns(
            AddColumns(
                GroupBy(
                    AddColumns(
                        Filter(Serviceberichte;Pruefer=User().FullName);
                        Monat; Month(Pruefdatum);
                        Jahr; Year(Pruefdatum)
                    );
                    Jahr;
                    Monat; 
                    ByDummy
                );
                Pruefzeitpunkt; Monat&"/"&Right(Jahr;2);
                '# per Pruefart';CountRows(ByDummy)
            );
            "Jahr";SortOrder.Ascending;
            "Monat";SortOrder.Ascending
        )
    )
    */
    //-----------------------------------------------------
    Set (_ShowLoadingHint,false);
    Set(FooterText,"");
);
