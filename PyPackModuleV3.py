def CloseWindow(window,rootwindow,hideroot=False):
    if hideroot:
        # To reveal it again:
        rootwindow.deiconify()
        # Hide it with .withdraw
        window.withdraw()
    else:
        # Destroy completely with .destroy
        window.destroy()
    return True

def HideWindow(window,hidewindow=True):
    if hidewindow:
        # Hide it with .withdraw
        window.withdraw()
    else:
        # To reveal it again:
        window.deiconify()
    return True

def SetWindowSize(window,size="800x1200",H=False,V=False):
    window.geometry(size)
    window.resizable(H,V)
    return True

def SetPackOption(varstroptmenu,Run_Button):
    if (Run_Button['state'] == 'disabled'):
        Run_Button.configure(state='normal')
    return varstroptmenu.get()

def CreateColumnHeadersInChannelSelectionFrame(treeview,treescroll,tvheight=16,configtvcol=False,TopWindow=False):
    #just add the headers and leave it
    treeview.delete(*treeview.get_children())
    treeview.configure(yscrollcommand=treescroll.set)
    if configtvcol==True :
        ConfigureColumnsInTree(treeview,tvheight,TopWindow)
    return True

def CreateColumnHeadersInShowPackFrame(treeview, treescroll):
    #just add the headers and leave it
    treeview.delete(*treeview.get_children())
    treeview.configure(yscrollcommand=treescroll.set)
    # Defining number of columns 
    #treeview["columns"] = ("1", "2", "3","4")
    #treeview["columns"] = ("one")
    # Defining heading 
    #treeview['show'] = 'headings'
    # Defining heading 
    treeview['height'] = 18
    treeview['selectmode'] ='browse'

    #treeview.column('0', width=40, minwidth=25, stretch='no')
    #treeview.column('one', width=500, minwidth=375, stretch='no')
    #treeview.column('two', width=225, minwidth=150, stretch='no')
    #treeview.column('three', width=200, minwidth=150, stretch='no')

    #treeview.heading('0', text='sl',anchor='w')
    #treeview.heading('one', text='broadcaster',anchor='w')
    #treeview.heading('two', text='pack name',anchor='w')
    #treeview.heading('three', text='channel name',anchor='w')
   
    return True

def ConfigureColumnsInTree(treeview,tvheight=16,TopWindow=False):
    
    if TopWindow:
        # Defining number of columns 
        #treeview["columns"] = ("1", "2", "3","4","5")
        treeview["columns"] = ("one", "two", "three","four","five","six")
        
        # Defining heading 
        treeview['show'] = 'headings'
        # Defining heading 
        treeview['height'] = tvheight
        treeview['selectmode'] ='browse'

        treeview.column('0',     width=40,  minwidth=25,  stretch='no')
        treeview.column('one',   width=40,  minwidth=25,  stretch='no')
        treeview.column('two',   width=200, minwidth=170, stretch='no')
        treeview.column('three', width=100, minwidth=50)
        treeview.column('four',  width=80,  minwidth=45,  stretch='no')
        treeview.column('five',  width=80,  minwidth=35,  stretch='no')
        treeview.column('six',   width=80,  minwidth=25,  stretch='no')
        #treeview.column('seven', width=40, minwidth=25, stretch='no')
        
        treeview.heading('0',     text='sl',      anchor='w')
        treeview.heading('one',   text='Typ',     anchor='w')
        treeview.heading('two',   text='channel_name',anchor='w')
        treeview.heading('three', text='cost',    anchor='w')
        treeview.heading('four',  text='genre',   anchor='w')
        treeview.heading('five',  text='lang',    anchor='w')
        treeview.heading('six',   text='choosen', anchor='w')
        #treeview.heading('seven', text='choosen',anchor='w')
    else:
        # Defining number of columns 
        #treeview["columns"] = ("1", "2", "3","4","5")
        treeview["columns"] = ("one", "two", "three","four")
        # Defining heading 
        treeview['show'] = 'headings'
        # Defining heading 
        treeview['height'] = tvheight
        treeview['selectmode'] ='browse'
        #treeview['selectmode'] ='none'

        treeview.column('0', width=40, minwidth=25, stretch='no')
        treeview.column('one', width=40, minwidth=25, stretch='no')
        treeview.column('two', width=240, minwidth=170, stretch='no')
        treeview.column('three', width=80, minwidth=50)
        treeview.column('four', width=100, minwidth=55, stretch='no')

        treeview.heading('0', text='sl',anchor='w')
        treeview.heading('one', text='sl',anchor='w')
        treeview.heading('two', text='channel_name',anchor='w')
        treeview.heading('three', text='cost',anchor='w')
        treeview.heading('four', text='choosen',anchor='w')
    return True

def ExportTreeViewDataInDF(treeview, pd, getalldata=False):
    #Get all channel selected + mandatory in a treeview_df
    # list of column names here
    df_colnames = ['slno','Channel_Name','Pack_Cost_Per_Month','Notes']
    tv_df = pd.DataFrame(columns=df_colnames)

    #print(treeview_df)
    for row in treeview.get_children():
        if getalldata == True:
            #print([treeview.item(row)["values"]])
            df_values = pd.DataFrame([treeview.item(row)["values"]],columns=df_colnames)
            #append the tree data to dataframe
            tv_df = tv_df.append(df_values,sort=True)
        else:
            if treeview.item(row)['values'][0] != "" and treeview.item(row)['values'][1] != "":
                # each row will come as a list under name "values" 
                #print([treeview.item(row)["values"]])
                df_values = pd.DataFrame([treeview.item(row)["values"]],columns=df_colnames)
                #append the tree data to dataframe
                tv_df = tv_df.append(df_values,sort=True)
    if getalldata == False:   
        tv_df['Choose_Yes_No'] = "Yes"
    tv_df['Pack_Cost_Per_Month'] = tv_df['Pack_Cost_Per_Month'].apply(pd.to_numeric)
    #print(tv_df.to_string())
    return tv_df

def RunPyRecommendationTool(PyPackMasterData, ServiceProvider,pd, mathObj, osObj, 
                            PackageDataPath, Run_Button, Report_Button, 
                            treeview1, treescroll1, treeview2, treescroll2,  
                            window, 
                            vartextmessage, varprogressmsg,varstrfilename,varblnfromfile,
                            MonthlyNetworkFee,
                            plt, np, MyDataResultFrame1):

    ##################################################################################################
    #Reading data from the excle file so change directory to location where you have kept the WishList 
    ##################################################################################################
    #print('blnflag value' + varblnfromfile.get())
        
    if MyDataResultFrame1["text"] != "Create Own Pack Data":
        #osObj.chdir(PyPackProgPath)
        Datafilename  = PackageDataPath
        ExcelFileName = Datafilename[Datafilename.rfind("\\")+1:] 
        #CurrDirectory = osObj.getcwd()

        #ServiceProviderName = ServiceProvider 'This is not required since we added a new param ServiceProvider to pass
        #ServiceProviderName = 'Airtel'
        SPPackfilename = PackageDataPath #osObj.path.join(CurrDirectory, ExcelFileName)
        PyPackProgPath = PackageDataPath #osObj.path.join(CurrDirectory, ExcelFileName)
        varstrfilename.set(Datafilename)
        
        #print("*****Execution Started****")
        #Read the Wsihlist and email id details for the user provided in MyPack and MyEmail worksheet
        df_myCurrentPack = ReadMyWishListPackData(PyPackMasterData, pd, mathObj,
                                                  Datafilename     = SPPackfilename,
                                                  MonthlyNetworkFee= MonthlyNetworkFee, 
                                                  FilterQuery      = 'KeepAll')
        #print(df_myCurrentPack['Send_EmailTo'])
        email_send = ','.join(df_myCurrentPack['Send_EmailTo'].unique()) # set the 'recipient_email'
        df_myCurrentPack = df_myCurrentPack.filter(items=['Channel_Name','Channel_Cost_Per_Month','Chosen'])
        #print("Reading my wishlist completed..." + Datafilename)
        varprogressmsg.set("Reading my wishlist completed from file.." + Datafilename)
        window.update()
    else:
        CurrDirectory   = osObj.getcwd()
        ExcelFileName   = "demoreportfile.txt"
        Datafilename    = "demoreportfile.txt"
        PackageDataPath = osObj.path.join(CurrDirectory, ExcelFileName)
        varstrfilename.set(Datafilename)
        
        #ServiceProviderName = ServiceProvider 'This is not required since we added a new param ServiceProvider to pass
        #ServiceProviderName = 'Airtel'
        SPPackfilename = PackageDataPath #osObj.path.join(CurrDirectory, ExcelFileName)
        PyPackProgPath = PackageDataPath #osObj.path.join(CurrDirectory, ExcelFileName)
    
        df_myCurrentPack  = ExportTreeViewDataInDF(treeview2, pd)
        df_myCurrentPack  = df_myCurrentPack.rename(columns = {'Choose_Yes_No'       : 'Chosen',
                                                               'Pack_Cost_Per_Month' : 'Channel_Cost_Per_Month'})
        df_myCurrentPack  = df_myCurrentPack.filter(items=['Channel_Name','Channel_Cost_Per_Month','Chosen'])
        varprogressmsg.set("Reading my wishlist completed from the treeview.")
        window.update()
        
    #Read all the pack data for the service provider from other sheets 
    df_OtherAvailablePacks = ReadAllServiceProviderPackData(PyPackMasterData, pd, SPPackfilename, ServiceProvider)
    #print("Reading all the pack details for the service provider completed..." + Datafilename)
    varprogressmsg.set("Reading all the pack details for the service provider completed..." + Datafilename)
    window.update()
    
    #Create an object of FPDF to be passed to the function
    #print("Start processing the recommendation report ..." + Datafilename)
    varprogressmsg.set("Start processing the recommendation report ..." + Datafilename)
    window.update()
    #print("Processing for Optimized Pack started....")
    varprogressmsg.set("Processing for Optimized Pack started....")
    window.update()
    df_optimizedPacks = PickOPtimizedCost(mathObj, osObj, pd, df_myCurrentPack, df_OtherAvailablePacks, varprogressmsg, window)

    #print("Processing for Optimized Pack completed....")
    varprogressmsg.set("Processing for Optimized Pack completed....")
    window.update()
    #print("Processing recommendation report for Optimized Pack....")
    varprogressmsg.set("Processing recommendation report for Optimized Pack....")
    window.update()
    
    pdf=object()
    #pdf=FPDF()
    fileHandler = object()
    #print('my pack')
    ProcessOptimizedPackToGenPyChannelPackageRecReportInTreeView(PyPackMasterData,
                                                                 treeview1, treescroll1, treeview2, treescroll2, 
                                                                 MyDataResultFrame1,  mathObj, pdf, fileHandler, osObj, pd,
                                                                 df_myCurrentPack, df_OtherAvailablePacks, df_optimizedPacks, 
                                                                 plt, np, Datafilename, PyPackProgPath, 
                                                                 ServiceProvider, False, True)
    #print("Recommendation report completed...."+ Datafilename)
    varprogressmsg.set("Recommendation report completed....")
    window.update()
    Report_Button.configure(state='normal')
    
    #print("*****Execution Completed****")
    varprogressmsg.set("*****Execution Completed****")
    window.update()
    
    #Destroy the RunRecommendation Button and the Message
    #Destroy the buttons and label if any from the frame  
    Run_Button.configure(state='disabled')   
    return True

def OpenFileDialog (MyDataResultFrame1, 
                    treeview2, treescroll2, treeview1, treescroll1, 
                    MyCreateReportbutton1, 
                    osObj,window,mathObj,Configurator,
                    vartextmessage, varprogressmsg, varblnfromfile, varstrfilename,
                    filedialog, pd, 
                    Run_Button,Report_Button):
    
    #get the current working directory
    cwd = osObj.getcwd()
    window.filename = filedialog.askopenfilename(initialdir= cwd,
                                                 title     = "Select an XL file",
                                                 filetypes = (("Excle file","*.xlsx"),("All files","*.*")))
    if window.filename == "":
        #set the from file flag as false
        varblnfromfile.set("No")
        
        #set the filename to blank
        varstrfilename.set("")
        vartextmessage.set("User cancelled selection of XL file... no data files selected.")
        window.update()
    else:
        #set the from file flag as true
        varblnfromfile.set("Yes")
        MyDataResultFrame1.config(text="Show Data From file")
        
        #store the filename in varstrfilename
        varstrfilename.set(window.filename)
        
        #show the filename in status
        vartextmessage.set(window.filename)
        window.update()
        
        #Read all the Service Provider's Pack Data the service provider name is part of the Datafilename sent across
        # Load spreadsheet
        xl = pd.ExcelFile(window.filename)
        
        #**Fetch & Store the data of MyPack in a new Data frame df_myCurrentPack and the Packs in df_OtherAvailablePacks Begin Code***
        blnInitializeOtherAvailablePacks = False
        # Get all the sheet names which has myCurrent Pack and other Packs to compare cost
        for sheetname in xl.sheet_names :
            # Load a sheet into a DataFrame by name: 
            if sheetname.lower() == 'mypack' :
                #set the custom columnnames that you want to use
                columnNames=['Channel_Name','Pack_Value_Monthly','Pack_Cost_Per_Month','Choose_Yes_No',
                             'Channel_Genre','Channel_Language','Channel_Type']
                if  not blnInitializeOtherAvailablePacks:    
                    dftemp = xl.parse(sheetname,names=columnNames)
                    blnInitializeOtherAvailablePacks = True
                    dftemp['Channel_Name'] = dftemp['Channel_Name'].str.strip()             #remove spaces
                    dftemp['Channel_Genre'] = dftemp['Channel_Genre'].str.strip()           #remove spaces
                    dftemp['Channel_Language'] = dftemp['Channel_Language'].str.strip()     #remove spaces
                    dftemp['Channel_Type'] = dftemp['Channel_Type'].str.strip()             #remove spaces
                    dftemp['Pack_Cost_Per_Month'] = dftemp['Pack_Cost_Per_Month']
                    
                    #Filter our channels which are selected as yes
                    df_mySelectedChannels     = dftemp.query("Choose_Yes_No=='Yes'")
                    if df_mySelectedChannels['Channel_Name'].count() >100 :
                        #Add the Fee for subsequent 25 channels
                        FeeSub25Chs = (df_mySelectedChannels['Channel_Name'].count()-1)
                        #print(FeeSub25Chs)
                        FeeSub25Chs = Calculate_NetworkFee_For25SubsequentChannels(mathObj, FeeSub25Chs)
                        #print(FeeSub25Chs)
                        
                        dftemp_chfee = pd.DataFrame([['Fee for Subsequent 25 Chs',FeeSub25Chs, FeeSub25Chs,'Yes',
                                                      'NA','NA','SD']],columns=columnNames)
                        dftemp       = dftemp.append(dftemp_chfee,sort=True)
                        
                    dftemp_chfee = pd.DataFrame([['Network Fee 0 to 100 Chs','153.40/Monthly',Configurator.MonthlyNetworkFee,'Yes',
                                                  'NA','NA','SD']], columns=columnNames)
                    dftemp       = dftemp.append(dftemp_chfee,sort=True)
                    
                    #Sort the data by channel cost
                    dftemp = dftemp.sort_values(by=['Pack_Cost_Per_Month','Channel_Name'],ascending=[False,True])
                    #filter for yes and cost>0
                    dftemp = dftemp.query("Choose_Yes_No=='Yes' and Pack_Cost_Per_Month >0")
                    #reset the index
                    dftemp = dftemp.reset_index(drop=True)
                    
                    CreateColumnHeadersInChannelSelectionFrame(treeview2,treescroll2,16)
                    blnInitializeOtherAvailablePacks = True
                    
                    #Display the choices made
                    int_rec_counter = 1
                    int_pack_cost = 0
                    for ind in dftemp.index:
                        #Display selection in Frame
                        if (dftemp['Pack_Cost_Per_Month'][ind]>0):
                            Notes = "Paid-" + dftemp['Channel_Type'][ind] 
                        elif(dftemp['Channel_Name'][ind][:2]=='dd'):
                            Notes = "Free=Mandatory" 
                        else:
                            Notes = "Free-" + dftemp['Channel_Type'][ind] 
                        treeview2.insert('','end',text=str(ind), 
                                         values=(str(ind+1), 
                                                 dftemp['Channel_Name'][ind],
                                                 str(dftemp['Pack_Cost_Per_Month'][ind]),
                                                 Notes))
                        int_pack_cost += dftemp['Pack_Cost_Per_Month'][ind]
                        int_rec_counter += 1
                
                #Display the pack cost in last row
                treeview2.insert('','end',text=str(int_rec_counter),
                                 values=("","Total Cost of Paid Channels (Rs)=",round(int_pack_cost,2),"",""))
                
                #Display the Run Pypack Recommendation Tool to get the suitable packs
                Run_Button.configure(state='normal')
                
                #Display the status bar to show the progress of PyPack recommendation tool
                varprogressmsg.set("Ready for execution....")
                window.update()
                
                #Create Column Headers PyPack Result
                CreateColumnHeadersInChannelSelectionFrame(treeview1,treescroll1,18,False,False)
                Report_Button.configure(state='disabled')
                break 
    return True

def CreateReport(tvcurrentpack, tvrecompack, pd, window, varprogressmsg, reportfilename="demoreportfile.txt"):
    
    #open a text file in read and write mode
    if reportfilename == '' or reportfilename == 'demoreportfile.txt':
        reportfilename = "demoreportfile.txt"
        txtfilename    = reportfilename
    else:
        txtfilename    = str(reportfilename).replace("xlsx","txt")
        txtfilename    = txtfilename.replace("xls","txt")
     
    #Open the report file  
    f = open(txtfilename,'w+')
    RecommendationText = "***********************************************************************************************"
    f.write( RecommendationText + "\n")
    RecommendationText = "                                   Py Recommendation Report                                    "
    f.write( RecommendationText + "\n")
    RecommendationText = "***********************************************************************************************"
    f.write( RecommendationText + "\n")
    #first print the current pack details
    top_rep_df = ExportTreeViewDataInDF(tvcurrentpack, pd, True)
    top_rep_df = top_rep_df.rename(columns = {'slno'                : 'Sl No',
                                              'Channel_Name'        : 'Channel Name',
                                              'Pack_Cost_Per_Month' : 'Cost_Per_Month',
                                              'Notes'               : 'Paid/Free'})
    top_rep_df = top_rep_df.filter(items=['Sl No','Channel Name','Cost_Per_Month','Paid/Free'])
    
    #for child in tvcurrentpack.get_children():
    #    listToStr = ' '.join([str(elem) for elem in tvcurrentpack.item(child)["values"]]) 
    #    f.write(listToStr + "\n")
    #    #print(tvcurrentpack.item(child)["values"])
    RecommendationText = top_rep_df.to_string()
    f.write( RecommendationText + "\n")
    RecommendationText = "***********************************************************************************************"
    f.write( RecommendationText + "\n")
    f.write( RecommendationText + "\n")
    
    #Next print the Recommendation from the Py Recommendation Tool
    packdet_rep_df = ExportTreeViewDataInDF(tvrecompack, pd, True)
    packdet_rep_df = packdet_rep_df.rename(columns = {'slno'                : 'Sl No',
                                                      'Channel_Name'        : 'Channel Name',
                                                      'Pack_Cost_Per_Month' : 'Pack Cost',
                                                      'Notes'               : 'Paid/Free'})
    packdet_rep_df = packdet_rep_df.filter(items=['Sl No','Channel Name','Pack Cost','Paid/Free'])
    #for child in tvrecompack.get_children():
    #    listToStr = ' '.join([str(elem) for elem in tvrecompack.item(child)["values"]]) 
    #    f.write(listToStr + "\n")
    #    #print(tvrecompack.item(child)["values"])
    #close the file
    RecommendationText = packdet_rep_df.to_string()
    f.write( RecommendationText + "\n")
    RecommendationText = "************************************************************************************************"
    f.write( RecommendationText + "\n")
    RecommendationText = "                                       Report End                                               "
    f.write( RecommendationText + "\n")
    RecommendationText = "************************************************************************************************"
    f.write( RecommendationText + "\n")
    f.close()
    #Show the message in status bar
    varprogressmsg.set("Report Created Successfully..." + txtfilename)
    window.update()
    return True

def FillPacksData(dftemp, treeview, treescroll):  
    
    #print(dftemp.to_string())
    #Clear the tree first
    treeview.delete(*treeview.get_children())
    treeview.configure(yscrollcommand=treescroll.set)
    #Set the row height and column width
    treeview['height']     = 16
    #treeview['width']     = 17
    treeview['selectmode'] ='extended'
    
    # Inserting items to the treeview 
    #Display the choices made
    if dftemp.empty :
        treeview.insert('','end',values=("**","There are no packs available for this Service Provider.","**","**"))
    else:
        
        # Inserting parent
        # Insert first all the Broadcaster details
        dfbroad = dftemp.filter(items=['Broadcaster_y'])
        dfbroad.drop_duplicates(inplace=True)
        int_broad_counter = 0
        for broadind in dfbroad.index:
            #Display selection in Frame 
            int_broad_counter += 1
            treeview.insert('',int_broad_counter,'item'+str(int_broad_counter), text=dfbroad['Broadcaster_y'][broadind])
            #treeview.insert('', '0', 'item1',text ='GeeksforGeeks')
            
            #set column width
            #treeview.column('', width=100, minwidth=55, stretch='no')
            
            dfpack = dftemp.query("Broadcaster_y=='" + dfbroad['Broadcaster_y'][broadind] + "'")
            dfpack = dfpack.filter(items=['Pack_Name2','Pack_Cost_Per_Month'])
            dfpack.drop_duplicates(inplace=True)
            int_pack_counter = int_broad_counter*1000
            for packind in dfpack.index:
                # Inserting child 
                int_pack_counter += 1
                treeview.insert('',int_pack_counter,'item'+str(int_pack_counter), 
                                text=dfpack['Pack_Name2'][packind] + " ["+ str(dfpack['Pack_Cost_Per_Month'][packind]) +"]")
                #treeview.insert('', '1', 'item2',text ='Computer Science')
                
                # Inserting more than one attribute of an item
                dfchan = dftemp.query("Pack_Name2=='" + dfpack['Pack_Name2'][packind] + "'")
                dfchan = dfchan.filter(items=['Channel_Name'])
                dfchan.drop_duplicates(inplace=True)
                
                for chanind in dfchan.index:
                    # Inserting more than one attribute of an item
                    treeview.insert('item'+str(int_pack_counter), 'end', 
                                    dfchan['Channel_Name'][chanind]+str(int_pack_counter),
                                    text = dfchan['Channel_Name'][chanind])  
                
                # Placing each child items in parent widget
                treeview.move('item'+str(int_pack_counter), 'item'+str(int_broad_counter), 'end') 
           
    return True  
def QueryChannelsData(dftemp, FilterQuery, treeview, treescroll, treeview1, treescroll1, pd, configtvcol=False):  
    
    #Filter out for channel type, language and genre
    dftemp = dftemp.query(FilterQuery)
    #configure this every time the query is refreshed
    CreateColumnHeadersInChannelSelectionFrame(treeview,treescroll,tvheight=18,configtvcol=configtvcol,TopWindow=True)
    if configtvcol:
        #configure this only for first time
        CreateColumnHeadersInChannelSelectionFrame(treeview1,treescroll1,tvheight=18,configtvcol=configtvcol,TopWindow=True)
    
    #Display the choices made
    int_rec_counter = 1
    int_pack_cost = 0
    for ind in dftemp.index:
        #Display selection in Frame 
        treeview.insert('','end',text=str(int_rec_counter), values=(dftemp['Channel_Type'][ind],
                                                        dftemp['Channel_Name'][ind],
                                                        str(dftemp['Pack_Cost_Per_Month'][ind]),
                                                        dftemp['Channel_Genre'][ind],
                                                        dftemp['Channel_Language'][ind],
                                                        dftemp['Choose_Yes_No'][ind]))
        int_pack_cost += dftemp['Pack_Cost_Per_Month'][ind]
        int_rec_counter += 1
    #Display message if none is available for selection in the given filter
    if int_rec_counter == 1:
        treeview.insert('','end',values=("**","There are no channels available for this filter.","**","**","**","**"))
    
    return True

def Calculate_NetworkFee_For25SubsequentChannels(mathObj, intNoOfChannels=1,FeeSub25ChsCost=20,FeeSub25ChsGSTPerc=18):
    '''
    Compute the Network fee for subsequent channels after 100 chs
    
    '''
    AdditionalChsafter100 = intNoOfChannels-100
    
    FeeSub25Chs = round(mathObj.ceil(AdditionalChsafter100/25)*(FeeSub25ChsCost+(FeeSub25ChsCost*FeeSub25ChsGSTPerc/100)),2)
    
    if AdditionalChsafter100>0 :
        return FeeSub25Chs
    else:
        return 0.0

def SendReport_EmailToReciepients(EMAILUSER,EMAILPWD,EMAILSENTTO,EMAILSUBJECT,
                                  RecTextReportPath,encoders,server,part,MIMEText,msg):
    '''
    send the recommendation report to recepient
    '''
    subject = 'PyPackage Recommender Report for ' + EMAILSUBJECT + ' Service Provider' # set the 'subject'
    msg['From'] = EMAILUSER
    msg['To'] = EMAILSENTTO
    #msg['Cc'] = EMAILUSER
    msg['Subject'] = subject

    body = 'Hi there, please find attached the recommendation report Generated from Python for your wishlist!'
    msg.attach(MIMEText(body,'plain'))

    #RecTextReportPath = 'Package_TataSky_NarayanaSwamy_ver4.txt' #code worked fine for text file kept in the same directory
    #RecTextReportPath = 'Package_TataSky_NarayanaSwamy_ver4.pdf'  #code worked fine for text file kept in the same directory

    filename= RecTextReportPath #set the 'filename'
    attachment  =open(filename,'rb')

    
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= "+filename)

    msg.attach(part)
    text = msg.as_string()
    
    server.starttls()
    server.login(EMAILUSER,EMAILPWD)

    server.sendmail(EMAILUSER,EMAILSENTTO,text)
    server.quit()
    return True

def ReadMyWishListPackData(PyPackMasterData, pd, mathObj,Datafilename='',MonthlyNetworkFee=153.40,Ch_Type='SD',FilterQuery='KeepYesSD'):
    '''
    Read the Wish list Data, Email id and Mandatory Channels
    '''
    # Load spreadsheet
    xl = pd.ExcelFile(Datafilename)
    Email_IDs ='NEP'
    #set the custom columnnames that you want to use while parsing
    columnNames=['Channel_Name','Channel_Value','Channel_Cost_Per_Month','Chosen','Genre','Language','Ch_Type']
        
    #**Fetch & Store the data of MyPack in a new Data frame df_myCurrentPack Begin Code***
    df_myCurrentPack = pd.DataFrame([['Network Fee 0 to 100 Chs','153.40/Monthly',MonthlyNetworkFee,'Yes','NA','NA','SD']],
                                    columns=columnNames)
    #'''
    #df_myCurrentPack = PyPackMasterData.GetAllChannelsData(pd,MonthlyNetworkFee)
    #'''
    #print("Count after initializeing {}".format(df_myCurrentPack['Channel_Name'].count()))
    # Get all the sheet names which has myCurrent Pack and other Packs to compare cost
    for sheetname in xl.sheet_names :
        # Load a sheet into a DataFrame by name: 
        if sheetname.lower() == 'mypack' :
            #df_myCurrentPack = xl.parse(sheetname,names=columnNames)
            dftemp = pd.DataFrame([],columns=columnNames)
            dftemp = xl.parse(sheetname,names=columnNames)
            df_myCurrentPack = df_myCurrentPack.append(dftemp,sort=True)
            #print("Count after adding wishlist {}".format(df_myCurrentPack['Channel_Name'].count()))
        elif sheetname.lower() == 'mymandatorychannels' :
            #df_myCurrentPack = xl.parse(sheetname,names=columnNames)
            dftemp = pd.DataFrame([],columns=columnNames)
            dftemp = xl.parse(sheetname,names=columnNames)
            dftemp['Choosen'] = 'Yes'   #Keep Yes since Mandatory
            dftemp['Ch_Type'] = 'SD'    #Keep SD since Mandatory
            #dftemp['Genre'] = 'GEC'    #Keep SD since Mandatory
            df_myCurrentPack = df_myCurrentPack.append(dftemp,sort=True)
            #print("Count after adding mand {}".format(df_myCurrentPack['Channel_Name'].count()))
        elif sheetname.lower() == 'myemail' :
            dftemp = xl.parse(sheetname,names=['EmailID'])
            dftemp['EmailID'] = dftemp['EmailID'].str.strip() #remove spaces
            dftemp = dftemp.drop_duplicates(keep='first') #remove duplicates
            Email_IDs = ';'.join(dftemp['EmailID'].tolist()) #add semicolon as email id seperator
    #**Fetch & Store the data of MyPack in a new Data frame df_myCurrentPack and the Packs in df_OtherAvailablePacks End Code***

    #********Clean up the data for spaces from the 2 dataframes Begin Code*******************************
    
    df_myCurrentPack['Channel_Name'] = df_myCurrentPack['Channel_Name'].str.strip() #remove spaces
    df_myCurrentPack['Channel_Name'] = df_myCurrentPack['Channel_Name'].str.lower() #convert to lower cases all the channels
    df_myCurrentPack['Genre'] = df_myCurrentPack['Genre'].str.strip()               #remove spaces
    df_myCurrentPack['Genre'] = df_myCurrentPack['Genre'].str.lower()               #convert to lower cases all the channels
    df_myCurrentPack['Language'] = df_myCurrentPack['Language'].str.strip()         #remove spaces
    df_myCurrentPack['Language'] = df_myCurrentPack['Language'].str.lower()         #convert to lower cases all the channels
    df_myCurrentPack['Ch_Type'] = df_myCurrentPack['Ch_Type'].str.strip()           #remove spaces
     
    #drop the duplicates
    df_myCurrentPack = df_myCurrentPack.drop_duplicates(keep='first')
    
    #add and set Send_EmailTO field 
    df_myCurrentPack['Send_EmailTo'] = Email_IDs
    df_myCurrentPack['Pack_Name1'] = 'Alacarte'
    df_myCurrentPack['Broadcaster_x'] = 'Alacarte'
    
   
    #Filter only Pack_Names which have cost greater than 0
    if FilterQuery =='KeepYesSD' :
        #df_myCurrentPack = df_myCurrentPack.query("Channel_Cost_Per_Month>0 and Chosen=='Yes' and Ch_Type=='" + Ch_Type +"'")
        df_myCurrentPack = df_myCurrentPack.query("Chosen=='Yes' and Ch_Type=='SD'")
    elif FilterQuery =='KeepYes' :
        df_myCurrentPack = df_myCurrentPack.query("Chosen=='Yes'")
    
    #Filter our channels which are selected as yes
    df_mySelectedChannels     = df_myCurrentPack.query("Chosen=='Yes'")
    #add susequent 25 chs cost if current selection is found>100
    if df_mySelectedChannels['Channel_Name'].count() >100 :
        #Add the Fee for subsequent 25 channels
        FeeSub25Chs = (df_mySelectedChannels['Channel_Name'].count()-1)
        #print(FeeSub25Chs)
        FeeSub25Chs = Calculate_NetworkFee_For25SubsequentChannels(mathObj, FeeSub25Chs)
        #print(FeeSub25Chs)
        columnNames = ['Channel_Name','Channel_Cost_Per_Month','Chosen','Send_EmailTo','Genre','Language','Ch_Type']
        dftemp = pd.DataFrame([['Fee for Subsequent 25 Chs',FeeSub25Chs,'Yes',Email_IDs,'NA','NA','SD']],
                              columns=columnNames)
        df_myCurrentPack = df_mySelectedChannels.append(dftemp,sort=True)
    
    #Sort the data by channel cost
    df_myCurrentPack = df_myCurrentPack.sort_values(by=['Channel_Cost_Per_Month','Channel_Name'],ascending=[False,True])
    #reset the index
    df_myCurrentPack = df_myCurrentPack.reset_index(drop=True)
    
    #Filter columns only Channel Name, Pack_Name and its monthly cost
    df_myCurrentPack = df_myCurrentPack.filter(items=['Channel_Name','Channel_Cost_Per_Month','Chosen',
                                                      'Send_EmailTo','Genre','Language','Ch_Type'])
    #********Clean up the data for spaces from the 2 dataframes End Code*******************************
    return df_myCurrentPack

def ReadAllServiceProviderPackData(PyPackMasterData, pd, Datafilename='',ServiceProvider='TataSky'):
    '''
    Read all the Service Provider's Pack Data the service provider name is part of the Datafilename sent across
    '''
    
    df_OtherAvailablePacks = PyPackMasterData.GetAllSPPackData(pd,ServiceProvider)
    df_OtherAvailablePacks = df_OtherAvailablePacks.sort_values(by=['Channel_Name'],ascending=[True])
    
    #********Clean up the data for spaces from the 2 dataframes End Code*******************************
    return df_OtherAvailablePacks


def PickOPtimizedCost(mathObj, osObj, pd, df_myCurrentPack, df_OtherAvailablePacks, varprogressmsg, window):
    '''
    Process the Wish list data against the service providers pack available and generate the recommendation report
    '''
    #Filter our channels which are selected as yes
    df_myCurrentPack     = df_myCurrentPack.query("Chosen=='Yes'")
    df_myCurrentPack     = df_myCurrentPack.reset_index(drop=True)
    
    #Initialise the Optimized pack and return in the end
    df_ReturnOptimizedPacks = pd.DataFrame([],columns=['PackOption', 'Pack_Name2', 'Pack_Cost_Per_Month', 'Channel_Cost', 
                                                       'Savings','Tot_MyPack_Cost', 'Tot_RecPack_Cost', 'Tot_NCF_Cost',
                                                       'Tot_Subs25Chs_Cost', 'Tot_Alacarte_Cost', 'Tot_Recommended_Cost', 
                                                       'Tot_RecPack_Savings'])
    #Initialize the packname as optioncounter
    InitializeOptionCounter =''
    
    #Copy the Other available packs in a dataframe for processing
    df_OtherAvlPacksforComparison = df_OtherAvailablePacks.copy(deep=True)
    
    df_mergePrimeCost = pd.merge(df_myCurrentPack,df_OtherAvlPacksforComparison,how='inner',on='Channel_Name')
    #df_mergedCostOrder['Savings'] = df_mergedCostOrder['Pack_Cost_Per_Month'] - df_mergedCostOrder[]
    df_mergePrimeCost = df_mergePrimeCost.groupby('Pack_Name2').agg({'Channel_Cost_Per_Month' : sum, 
                                                                       'Pack_Cost_Per_Month':  max, 
                                                                       'Broadcaster_y': "count"}) 
    df_mergePrimeCost = df_mergePrimeCost.rename(columns = {'Broadcaster_y':'Pack_Count',
                                                              'Channel_Cost_Per_Month' : 'Channel_Cost',
                                                              'Pack_Cost_Per_Month':'Pack_Cost'})
    df_mergePrimeCost = df_mergePrimeCost[df_mergePrimeCost['Pack_Cost'] <= df_mergePrimeCost['Channel_Cost']]
    df_mergePrimeCost['PrimeSavings'] = (df_mergePrimeCost['Channel_Cost']-df_mergePrimeCost['Pack_Cost']).round(2)
    
    df_mergePrimeCost = df_mergePrimeCost.sort_values(by=['PrimeSavings','Pack_Count'],ascending=[False,True])
    
    #Get only those costs which have same savings
    #print(df_mergePrimeCost)
    if df_mergePrimeCost.empty:
        df_optimizedPackDetails = pd.DataFrame([],columns=['Pack_Name2','Pack_Cost_Per_Month','Channel_Cost','Savings','Iteration'])
        #Return and empty Pack list since there are no optimized packs available 
        df_ReturnOptimizedPacks = df_ReturnOptimizedPacks.append(df_optimizedPackDetails,sort=True)
        return df_ReturnOptimizedPacks
    else:
        
        df_mergePacks  = pd.merge(df_myCurrentPack,df_OtherAvlPacksforComparison,how='inner',on='Channel_Name')
        df_mergePacks  = df_mergePacks.groupby('Pack_Name2', as_index=False).agg({'Channel_Cost_Per_Month' : sum,
                                                                  'Pack_Cost_Per_Month':  max,
                                                                  'Broadcaster_y': "count"}) 
        df_mergePacks  = df_mergePacks.rename(columns = {'Broadcaster_y':'Pack_Count',
                                                         'Channel_Cost_Per_Month' : 'Channel_Cost',
                                                         'Pack_Cost_Per_Month':'Pack_Cost'})
        df_mergePacks = df_mergePacks[df_mergePacks['Pack_Cost'] <= df_mergePacks['Channel_Cost']]
        df_mergePacks['PrimeSavings'] = (df_mergePacks['Channel_Cost']-df_mergePacks['Pack_Cost']).round(2)
        #print(df_mergePacks) 

        df_grpgetcount = df_mergePacks.groupby(['PrimeSavings','Pack_Cost'], as_index=False).count()
        #print(df_grpgetcount)
        df_dupPackSamecosts = df_grpgetcount[df_grpgetcount['Pack_Count']>1]
        #print(df_dupPackSamecosts)
        saving_ids = df_dupPackSamecosts['PrimeSavings']

        #df_PackSameSavingstemp = df_PackSameSavingstemp[df_PackSameSavingstemp['Pack_Count']>1]

        #saving_ids = df_PackSameSavingstemp.groups.keys()

        #saving_ids = df_PackSameSavingstemp["PrimeSavings"]
        print("duplicate pack cost savings" + str(len(saving_ids)))

        if df_dupPackSamecosts.empty:
            #PrimeCostSavingsPackNameList = df_mergePrimeCost.index.tolist()
            IntListSavings = 1
            myList = []
            for items  in range(0,IntListSavings,1):
                myList.append("Pack no.." + str(items+1))
            PrimeCostSavingsPackNameList = myList
        else:
            IntListSavings = len(saving_ids)
            myList = []
            for items  in range(0,IntListSavings,1):
                myList.append("Pack no.." + str(items+1))

            df_PackSameSavings = df_mergePacks[df_mergePacks['PrimeSavings'].isin(saving_ids)]
            PrimeCostSavingsPackNameList = myList

        #print(PrimeCostSavingsPackNameList)

        #IntListSavings = len(df_PackSameSavings)
        #print(IntListSavings)

        #for primepack in PrimeCostSavin1gsPackNameList:
        for x in range(0,IntListSavings,1):

            #Initialize all the other dataframes 
            if len(PrimeCostSavingsPackNameList) == 0 : 
                primepack = 'NothingToChose'
            else:
                primepack = PrimeCostSavingsPackNameList[x]

            InitializeOptionCounter = primepack
            print("checking pack - {}...".format(InitializeOptionCounter))
            varprogressmsg.set("checking pack - {}...".format(InitializeOptionCounter))
            window.update()
            #Use this data frame to store all the selected packs which offer savings
            df_optimizedPackDetails = pd.DataFrame([],columns=['Pack_Name2','Pack_Cost_Per_Month','Channel_Cost',
                                                               'Savings','Iteration'])
            df_optimizedPackCostDetails = pd.DataFrame([],columns=['PackOption','Tot_MyPack_Cost','Tot_RecPack_Cost',
                                                                   'Tot_NCF_Cost','Tot_Subs25Chs_Cost','Tot_Alacarte_Cost',
                                                                   'Tot_Recommended_Cost','Tot_RecPack_Savings','Iteration'])
            #Initialise the cost of selected packs
            CostofMypack = 0.0
            df_selectedpackNames = pd.DataFrame([],columns=['Pack_Name2','Pack_Cost_Per_Month','Channel_Cost','Savings','Iteration'])

            #*Filter data from myCurrentPack to get only channels with cost > 0 and get total cost of pack with choosen= yes BEGIN Code**   

            #Store the Original channels list in a Dataframe for removing them in each iteration
            df_remaining = df_myCurrentPack.sort_values(by=['Channel_Name'] ,axis=0)

            #***Filter data from myCurrentPack to get only channels with cost > 0 and get total cost of pack with choosen= yes END Code**

            #******* Pick only those packs where we see any cost benefit and arrange them from max to mix selection order BEGIN CODE**
            #***This code is required to reduce iterations and ordering is CRITICAL to allow the projection from max selection to min

            df_mergedCostOrder = pd.merge(df_myCurrentPack,df_OtherAvlPacksforComparison,how='inner',on='Channel_Name')
            #df_mergedCostOrder['Savings'] = df_mergedCostOrder['Pack_Cost_Per_Month'] - df_mergedCostOrder[]
            df_mergedCostOrder = df_mergedCostOrder.groupby('Pack_Name2').agg({'Channel_Cost_Per_Month' : sum, 
                                                                               'Pack_Cost_Per_Month':  max, 
                                                                               'Broadcaster_y': "count"}) 
            df_mergedCostOrder = df_mergedCostOrder.rename(columns = {'Broadcaster_y':'Pack_Count',
                                                                      'Channel_Cost_Per_Month' : 'Channel_Cost',
                                                                      'Pack_Cost_Per_Month':'Pack_Cost'})
            df_mergedCostOrder = df_mergedCostOrder[df_mergedCostOrder['Pack_Cost'] <= df_mergedCostOrder['Channel_Cost']]
            df_mergedCostOrder['Savings'] = (df_mergedCostOrder['Channel_Cost']-df_mergedCostOrder['Pack_Cost']).round(2)
            #print(df_mergedCostOrder)

            #fixed the sort column and mad ascending for pack count 
            #df_mergedCostOrder = df_mergedCostOrder.sort_values(by=['Pack_Count','Savings'],ascending=[False,False])
            df_mergedCostOrder = df_mergedCostOrder.sort_values(by=['Savings','Pack_Count'],ascending=[False,True])

            #print (df_mergedCostOrder)
            #Get the Packnames into a list and then use this for iterating through base and other pack comparison
            CostSavingsPackNameList = df_mergedCostOrder.index.tolist()

            #******* Pick only those packs where we see any cost benefit and arrange them from max to mix selection order End CODE**

            #***********Compare all the packs and fetch the optimum cost of the pack Outer Iteration Begin*************************
            for packname in CostSavingsPackNameList:

                df_CompPack = df_OtherAvlPacksforComparison[df_OtherAvlPacksforComparison['Pack_Name2']==packname]
                df_CompPack = df_CompPack.filter(items=['Channel_Name','Pack_Cost_Per_Month','Pack_Name2','Savings'])
                df_BaseMergePack = pd.merge(df_remaining,df_CompPack,how='inner',on='Channel_Name')

                CostSaving_forBasePack_per_month = df_BaseMergePack['Channel_Cost_Per_Month'].sum() -\
                                                    df_BaseMergePack['Pack_Cost_Per_Month'].max()

                df_BaseMergePack['Channel_Cost'] = df_BaseMergePack['Channel_Cost_Per_Month'].sum()
                df_BaseMergePack['Savings']      = CostSaving_forBasePack_per_month
                df_CompPack['Savings'] = CostSaving_forBasePack_per_month

                df_BasePack = df_BaseMergePack.filter(items=['Channel_Name','Channel_Cost_Per_Month'])

                blnLessCostPackfound = False
                #**************compare only when the CostSaving_forBasePack_per_month is greater than 0 begin code*******
                if (CostSaving_forBasePack_per_month>=0) :

                    #********compare base with other packs through iteration Begin**************************************************
                    for comparepackname in CostSavingsPackNameList:
                        if comparepackname != packname:

                            #compare the other packs cost with this cost
                            df_OtherPack = df_OtherAvlPacksforComparison[df_OtherAvlPacksforComparison['Pack_Name2']
                                                                         ==comparepackname]
                            df_OtherPack = df_OtherPack.filter(items=['Channel_Name','Pack_Cost_Per_Month',
                                                                      'Pack_Name2','Savings'])
                            df_mergedComp = pd.merge(df_BasePack,df_OtherPack,how='inner',on='Channel_Name')

                            CostSaving_forOtherPack_per_month =  df_mergedComp['Channel_Cost_Per_Month'].sum() -\
                                                                 df_mergedComp['Pack_Cost_Per_Month'].max()

                            df_mergedComp['Channel_Cost'] = df_mergedComp['Channel_Cost_Per_Month'].sum()
                            df_mergedComp['Savings'] = CostSaving_forOtherPack_per_month.round(2)
                            df_OtherPack['Savings'] = CostSaving_forOtherPack_per_month.round(2)
                            if (CostSaving_forOtherPack_per_month > CostSaving_forBasePack_per_month) :
                                #select this pack as this is offering better cost and discard the other
                                #Store the packname and costs in another temp dataframe
                                dftemp              = df_mergedComp.filter(items=['Pack_Name2','Pack_Cost_Per_Month',
                                                                                  'Channel_Cost','Savings'])
                                dftemp['Iteration'] = x
                                dftemp              = dftemp.drop_duplicates()
                                blnLessCostPackfound = True
                    #****************compare base with other packs through iteration End**************************************************

                    if blnLessCostPackfound :
                        #Store the dftemp packname into df_selectedpackNames data frame as that was the best among the comparison
                        df_selectedpackNames = df_selectedpackNames.append(dftemp,sort=True)

                        df_opttemp = dftemp.filter(items=['Pack_Name2'])
                        df_opttemp['PackOption'] = InitializeOptionCounter
                        df_opttemp = df_opttemp.drop(['Pack_Name2'],axis=1)
                        df_optimizedPackDetails  = df_optimizedPackDetails.append(pd.concat([dftemp,df_opttemp],
                                                                                            axis=1,sort=True),sort=True)
                        #print(df_optimizedPackDetails)

                        #Remove only those channels where a selection is made and retain others so that duplicate selection is avoided 
                        #and other channel gets a chance for selection during the succesive iterations
                        SelectedPackName = dftemp['Pack_Name2'].unique()
                        dftemp = df_OtherAvlPacksforComparison[df_OtherAvlPacksforComparison['Pack_Name2']==SelectedPackName[0]]
                        dftemp = dftemp.filter(items=['Channel_Name'])
                        dftemp = dftemp.drop_duplicates()
                        df_remaining = df_remaining[~df_remaining['Channel_Name'].isin(dftemp['Channel_Name'])]

                    else:
                        #Store the Base pack if the Savings is more than 0 as that is the best into df_selectedpackNames data frame
                        dftemp               = df_BaseMergePack[df_BaseMergePack['Savings']>0]
                        dftemp               = df_BaseMergePack.filter(items=['Pack_Name2','Pack_Cost_Per_Month',
                                                                              'Channel_Cost','Savings'])
                        dftemp['Iteration']  = x
                        dftemp               = dftemp.drop_duplicates()
                        df_selectedpackNames = df_selectedpackNames.append(dftemp,sort=True)

                        df_opttemp = dftemp.filter(items=['Pack_Name2'])                
                        df_opttemp['PackOption'] = InitializeOptionCounter
                        df_opttemp = df_opttemp.drop(['Pack_Name2'],axis=1)

                        df_optimizedPackDetails = df_optimizedPackDetails.append(pd.concat([dftemp,df_opttemp],
                                                                                           axis=1,sort=True),sort=True)
                        #df_optimizedPackDetails = df_optimizedPackDetails.append(dftemp)
                        #print(df_optimizedPackDetails)

                        #Remove the channels when a selection is made so as to avoid duplicate selection during the succesive iterations
                        df_remaining = df_remaining[~df_remaining['Channel_Name'].isin(df_BasePack['Channel_Name'])]
                #else:
                    #ignore the negative savings and we dont remove any channels but use them for next pack comparison
                    ##print("Packname {} has cost = {}".format(packname,CostSaving_forBasePack_per_month))

                #**************compare only when the CostSaving_forBasePack_per_month is greater than 0 End code*******

            #***********Compare all the packs and fetch the optimum cost of the pack Outer Iteration End*************************


            #********************Make the recommendtions BEGIN******************************************************************
            #Get all the channels from the selected packs in a df_SelectedChannelsList dataframe to see what additional channels user gets

            df_SelectedChannelsList = df_OtherAvlPacksforComparison[df_OtherAvlPacksforComparison['Pack_Name2'].
                                                             isin(df_selectedpackNames['Pack_Name2'])].filter(items=['Channel_Name'])
            ##print(df_SelectedChannelsList)
            #******* First print the current Pack details BEGIN CODE**************************************************************

            #compute the cost of current pack
            CostofMypack = df_myCurrentPack['Channel_Cost_Per_Month'].sum()
            #print("Cost of my pack =" + str(CostofMypack))
            #******* First print the current Pack details END CODE**************************************************************

            if df_selectedpackNames['Pack_Name2'].count()!=0 :
                #********Next show the selected pack with savings that it will make Begin code***************************************
                CostofSelectedPacks = 0
                RecommendationText = ""

                #****************************Start the selected package details in the SECOND PAGE  Begin Code **********************

                df_selectedpackNames = df_selectedpackNames.reset_index(drop=True)
                df_optimizedPackDetails = df_optimizedPackDetails.reset_index(drop=True)

                CostofRecommendedPacks = df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2)
                CostofSelectedPacks += df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2)
                CostofNCF      = 0.0
                CostofSub25Chs = 0.0
                CostofAlac     = 0.0

                #********Next show the selected pack with savings that it will make End code************************************

                #********Next show the individual channels that could not go in with any pack Begin code*************************
                #********Start writing individual channels into FOURTH PAGE Begin Code*********************************
                #Add alcarte details in a new page - page no 4
                #Get the remaining by removing the selected channels from selectepackName
                df_remaining = df_myCurrentPack[~df_myCurrentPack['Channel_Name'].isin(df_SelectedChannelsList['Channel_Name'])]


                #Get the additional channels from the selected packs that were not part of the current list
                #Lets use the set theory and use DIFFERENCE operation to achieve the results 

                #Get the selected channels in a new set from the dataframe of selectec packs
                SelectedChannels=set(df_SelectedChannelsList['Channel_Name'])
                #print(SelectedChannels.to_string())
                #Get the original channels from current pack list in another set
                OriginalChannels=set(df_myCurrentPack['Channel_Name'])
                #print(OriginalChannels.to_string())
                #AdditionalChannels = SelectedChannels - OriginalChannels
                AdditionalChannels = SelectedChannels.difference(OriginalChannels)
                #print(AdditionalChannels.to_string())

                AdditionChannelsCount = len(AdditionalChannels)
                BasePackCount = (df_myCurrentPack['Channel_Name'].count()-1)
                if df_remaining['Channel_Name'].count()==1 :
                    if (BasePackCount + AdditionChannelsCount) <=100 :
                        RecommendationText = "Add the Network Fee :"
                        FeeSub25Chs = 0
                    elif (BasePackCount + AdditionChannelsCount) > 100 :
                        RecommendationText = "Add the Network Fee + Fee for Subsequent 25 channels:"
                        #add the Fee for subsequent 25 channels
                        FeeSub25Chs = (df_myCurrentPack['Channel_Name'].count() + len(AdditionalChannels))
                        FeeSub25Chs = Calculate_NetworkFee_For25SubsequentChannels(mathObj, FeeSub25Chs)
                        dftemp = pd.DataFrame([['Fee for Subsequent 25 Chs',FeeSub25Chs]],
                                              columns=['Channel_Name','Channel_Cost_Per_Month'])
                        df_remaining = df_remaining.append(dftemp,ignore_index = True,sort=True)
                else:
                    if (BasePackCount + AdditionChannelsCount) <=100 :
                        RecommendationText = "Add the Network Fee plus below channels as Alacarte:"
                        FeeSub25Chs = 0
                    elif (BasePackCount + AdditionChannelsCount) > 100 :
                        RecommendationText = "Add the Network Fee + Fee for Subsequent 25 channels plus below channels as Alacarte:"
                        #add the Fee for subsequent 25 channels
                        FeeSub25Chs = (BasePackCount + AdditionChannelsCount)
                        FeeSub25Chs = Calculate_NetworkFee_For25SubsequentChannels(mathObj, FeeSub25Chs)
                        dftemp = pd.DataFrame([['Fee for Subsequent 25 Chs',FeeSub25Chs]],
                                              columns=['Channel_Name','Channel_Cost_Per_Month'])
                        df_remaining = df_remaining.append(dftemp,ignore_index = True,sort=True)

                df_remaining = df_remaining.filter(items=['Channel_Name','Channel_Cost_Per_Month'])
                df_remaining = df_remaining.sort_values(by=['Channel_Cost_Per_Month','Channel_Name'],ascending=[False,True])
                df_remaining = df_remaining.drop_duplicates(keep='first')
                df_remaining = df_remaining.reset_index(drop=True)

               #compute the cost of alacarte + network + Fee for sub 25 chs 
                CostofSelectedPacks += df_remaining['Channel_Cost_Per_Month'].sum().round(2)
                #print("cost of selected packs"+str(CostofSelectedPacks))
                #********Next show the individual channels that could not go in with any pack End code*************************
                #*Finally show the Total cost of new pack and its Savings & print the additional channels user will benefit Begin code*

                df_opttempdf = pd.DataFrame([[InitializeOptionCounter,CostofMypack,CostofRecommendedPacks,CostofNCF,
                                             CostofSub25Chs, CostofAlac, (CostofRecommendedPacks+CostofNCF+CostofSub25Chs+CostofAlac), 
                                             #(CostofMypack - CostofRecommendedPacks-CostofNCF-CostofSub25Chs-CostofAlac)
                                              (CostofMypack -CostofSelectedPacks), x, len(AdditionalChannels)]],
                                            columns=['PackOption','Tot_MyPack_Cost','Tot_RecPack_Cost','Tot_NCF_Cost',
                                                     'Tot_Subs25Chs_Cost','Tot_Alacarte_Cost','Tot_Recommended_Cost',
                                                     'Tot_RecPack_Savings','Iteration','AdditionalChannels'])            
                df_opttempdf = pd.merge(df_opttempdf,df_optimizedPackDetails,how="inner",on="PackOption")
                df_opttempdf = df_opttempdf.filter(items=['PackOption','Tot_MyPack_Cost','Tot_RecPack_Cost','Tot_NCF_Cost',
                                                          'Tot_Subs25Chs_Cost','Tot_Alacarte_Cost','Tot_Recommended_Cost',
                                                          'Tot_RecPack_Savings','Iteration','AdditionalChannels'])
                df_optimizedPackCostDetails = df_optimizedPackCostDetails.append(df_opttempdf,sort=True)

                df_optimizedPackDetails = pd.merge(df_optimizedPackDetails,df_optimizedPackCostDetails,how="inner",on="PackOption")
                # dropping duplicate values 
                df_optimizedPackDetails = df_optimizedPackDetails.drop_duplicates(keep='first')
                df_optimizedPackDetails = df_optimizedPackDetails.reset_index(drop=True)
                #print(df_optimizedPackDetails.to_string())

                df_ReturnOptimizedPacks = df_ReturnOptimizedPacks.append(df_optimizedPackDetails,sort=True)

            #*********** Recommendation pack stored in optimized pack*******************

            #*************************CRITICAL CODE for filetering already compared PRIME Pack Name BEGIN********************
            df_OtherAvlPacksforComparison = df_OtherAvlPacksforComparison[df_OtherAvlPacksforComparison['Pack_Name2'] != 
                                                                          InitializeOptionCounter.strip()]

        #*************************CRITICAL CODE for filetering already compared PRIME Pack Name END********************
        #Filter and return only the Max Savings Record
        MaxSavings    = df_ReturnOptimizedPacks["Tot_RecPack_Savings"].max()
        MaxChannels   = df_ReturnOptimizedPacks["AdditionalChannels"].max()
        MinChannels   = df_ReturnOptimizedPacks["AdditionalChannels"].min()


        #print("Max Savings=" + str(MaxSavings))

        #Pick the MaxSavings
        df_ReturnOptimizedPacks = df_ReturnOptimizedPacks[df_ReturnOptimizedPacks["Tot_RecPack_Savings"]==MaxSavings]
        #Pick the iteration which offers max channels if both are equal then pick any iteration.

        df_ReturnOptimizedPacks = df_ReturnOptimizedPacks[df_ReturnOptimizedPacks["AdditionalChannels"]==MaxChannels]
        MinIteration  = df_ReturnOptimizedPacks["Iteration_x"].min()
        MaxIteration  = df_ReturnOptimizedPacks["Iteration_x"].max()
        if MaxIteration != MinIteration:
            df_ReturnOptimizedPacks = df_ReturnOptimizedPacks[df_ReturnOptimizedPacks["Iteration_x"]==MaxIteration]
        elif MaxIteration == MinIteration:
            df_ReturnOptimizedPacks = df_ReturnOptimizedPacks[df_ReturnOptimizedPacks["Iteration_x"]==MinIteration]
        
        #reset the index
        df_ReturnOptimizedPacks = df_ReturnOptimizedPacks.reset_index(drop=True)
        #print(df_ReturnOptimizedPacks.to_string())
        return df_ReturnOptimizedPacks
        #********************Make the recommendations END******************************************************************

def ProcessOptimizedPackToGenPyChannelPackageRecReportInTreeView(PyPackMasterData,
                                                                 treeview1, treescroll1, treeview2, treescroll2, 
                                                                 MyDataResultFrame1,  mathObj, pdf, fileHandler, osObj, 
                                                                 pd, 
                                                                 df_myCurrentPack, df_OtherAvailablePacks, df_OptimizedPack, 
                                                                 plt, np, 
                                                                 Datafilename='', PyRecReportPath='', 
                                                                 ServiceProvider='Tata Sky', 
                                                                 blnWriteToText=False, blnWriteToPDF=True):
    '''
    Process the Wish list data against the service providers pack available and generate the recommendation report
    '''
    df_AllChannelsList   = PyPackMasterData.GetAllChannelsData(pd,MonthlyNetworkFee=153.40,MandatoryOnly='All')
    df_AllChannelsList   = df_AllChannelsList.rename(columns = {'Choose_Yes_No'       : 'Chosen',
                                                                'Pack_Cost_Per_Month' : 'Channel_Cost_Per_Month'})
    #***Filter data from myCurrentPack to get only channels with cost > 0 and get total cost of pack with choosen= yes BEGIN Code**
    #Filter our channels which are selected as yes
    df_myCurrentPack     = df_myCurrentPack.query("Chosen=='Yes'")
    df_myCurrentPack     = df_myCurrentPack.reset_index(drop=True)
    
    #Initialise the cost of selected packs
    CostofMypack = 0.0
    #df_selectedpackNames = pd.DataFrame([],columns=['Pack_Name2','Pack_Cost_Per_Month','Channel_Cost','Savings'])
    df_selectedpackNames = df_OptimizedPack.filter(items=['Pack_Name2','Pack_Cost_Per_Month','Channel_Cost','Savings'])
    #Store the Original channels list in a Dataframe for removing them in each iteration
    df_remaining = df_myCurrentPack.sort_values(by=['Channel_Name'] ,axis=0)
        
    #compute the cost of current pack
    CostofMypack = df_myCurrentPack['Channel_Cost_Per_Month'].sum()

    #********************Make the recommendations BEGIN******************************************************************
    #Get all the channels from the selected packs in a df_SelectedChannelsList dataframe to see what additional channels user gets
    #df_OtherAvailablePacks = ReadAllServiceProviderPackData(pd,Datafilename)
    df_SelectedChannelsList = df_OtherAvailablePacks[df_OtherAvailablePacks['Pack_Name2'].
                                                     isin(df_OptimizedPack['Pack_Name2'])].filter(items=['Channel_Name'])
    #******************Adding the current pack Details***************Page 1 BEGIN Code **********************
    #blnPrintHeader = False
    #for ind in df_myCurrentPack.index:
    #    if not blnPrintHeader: 
    #                  
    #        CreateColumnHeadersInChannelSelectionFrame(treeview1,treescroll1,18)
    #        CreateColumnHeadersInChannelSelectionFrame(treeview2,treescroll2,16)
    #        blnPrintHeader = True
    #        
    #    #Display selection in Frame
    #    if (df_myCurrentPack['Channel_Cost_Per_Month'][ind]>0):
    #        Notes = "Paid"
    #    elif (df_myCurrentPack['Channel_Name'][ind][:2]=='dd'):
    #        Notes = "Free=Mandatory"
    #    else:
    #        Notes = "Free"
    #    parentnode = treeview2.insert('',(ind+1),str(ind+1),text="C"+str(ind+1),
    #                                  values=(str(ind+1),
    #                                          df_myCurrentPack['Channel_Name'][ind],
    #                                          str(df_myCurrentPack['Channel_Cost_Per_Month'][ind]),
    #                                          Notes))
    #latestindex = df_myCurrentPack.index.max()+2
    #parentnode = treeview2.insert('',(latestindex),str(latestindex),text="C"+str(latestindex),
    #                 values=("","Cost of Your current Pack (Rs)=",round(CostofMypack,2),""))
    
    #Don't change the caption as its being used as a flag now to control the process flow for data from file flag selection
    #MyDataResultFrame1.config(text="Your Current Pack's Cost details")
    
    
    #******************Adding the Recommended Pack Details****************************Page 2 Begin Code ******************
    #treeview1.insert('','end',values=("","Recommendation Report for " + ServiceProvider,"",""))
    #treeview1.insert('','0','item0',text= '0',values=("","Recommendation Report for " + ServiceProvider,"",""))
    if df_selectedpackNames['Pack_Name2'].count()==0 :
        
        #********Case when NO selected pack available-- Begin code***************************************
        treeview1.insert('','end',values=("","There are NO PACKS available that offer any cost savings","",""))
        treeview1.insert('','end',values=("","Hence, Recommendation is that you buy them as ALACARTE.","",""))
        #********Case when NO selected pack available-- End code***************************************
    else:
        
        #********Next show the selected pack with savings that it will make Begin code***************************************
        CostofSelectedPacks = 0
        
        #Adding recommended package in page no 2
        #Inserting Root Parent Node
        treeview1.insert('','1','item1',text= '1',values=("","The Recommendation is to add the below package(s)","",""))
         
        df_selectedpackNames = df_selectedpackNames.reset_index(drop=True)
        
        blnPrintHeader = False
        intlastindex = 0
        treeview1.tag_configure("evenrow",background='orange',foreground='blue')
        treeview1.tag_configure("oddrow",background='black',foreground='white')
        treeview1.tag_configure("commonrow",background='green',foreground='white')
        for ind in df_selectedpackNames.index:
            if (df_selectedpackNames['Savings'][ind]>0):
                Notes = "Savings=" + str(round(df_selectedpackNames['Savings'][ind],2))
            else:
                Notes = "Savings=0"
            
            #Inserting first level Parent Node
            parentnode = treeview1.insert('',(ind+1),str(ind+1),text="P"+str(ind+1),tags=('evenrow',),
                                          values=(str(ind+1),
                                                  df_selectedpackNames['Pack_Name2'][ind],
                                                  str(df_selectedpackNames['Pack_Cost_Per_Month'][ind]),
                                                  Notes))
            '''
            Get all the channel for that pack and highlight in green all that is present in the user selection
            '''
            #Insert child as channels for that pack
            df_childchannelNames =  df_OtherAvailablePacks.query("Pack_Name2=='" + df_selectedpackNames['Pack_Name2'][ind] + "'")
            df_mergetemp         = pd.merge(df_childchannelNames,df_AllChannelsList,on='Channel_Name',how='inner')
            #df_mergetemp         = pd.merge(df_childchannelNames,df_AllChannelsList,on='Channel_Name',how='left')
            
            df_mergetemp = df_mergetemp.reset_index(drop=True)
            intlastindex += df_mergetemp.index.max()
            
            for childind in df_mergetemp.index:
                # check if the channel from pack is user selection also
                df = df_myCurrentPack.query("Channel_Name=='" + str(df_mergetemp['Channel_Name'][childind])+ "'")
                if mathObj.isnan(df.index.max()): 
                    treeview1.insert(parentnode,'end',str(ind+1)+str(childind+1),text=(str(ind+1)+str(childind+1)), tags=('oddrow',),
                                 values=("   "+str(ind+1) + "." + str(childind+1),
                                         "     " + df_mergetemp['Channel_Name'][childind],
                                         df_mergetemp['Channel_Cost_Per_Month'][childind],
                                         "Pack" + str(ind+1)+"Child"))
                else:
                    treeview1.insert(parentnode,'end',str(ind+1)+str(childind+1),text=(str(ind+1)+str(childind+1)), tags=('commonrow',),
                                 values=("   "+str(ind+1) + "." + str(childind+1),
                                         "     " + df_mergetemp['Channel_Name'][childind],
                                         df_mergetemp['Channel_Cost_Per_Month'][childind],
                                         "Pack" + str(ind+1)+"Child"))

                
             
        CostofSelectedPacks = df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2)
        intlastindex += df_selectedpackNames.index.max() + 2
        treeview1.insert('','end',values=("","Cost of the Recommended pack(Rs)=",round(CostofSelectedPacks,2),""))
        #parentnode = treeview1.insert('',(intlastindex),str(intlastindex),text="P"+str(intlastindex),
        #                              values=("","Cost of the Recommended pack(Rs)=",round(CostofSelectedPacks,2),"",""))
         #********Next show the selected pack with savings that it will make End code***************************************
        
    #******************Adding the Recommended Pack Details****************************Page 2 End Code ******************
       
    #******************Adding the Individual channels(Alacarte Details)******************Page 4 Begin Code ******************
    
        #********Start writing individual channels into FOURTH PAGE*********************************
        #Add alcarte details in a new page - page no 4
        #Get the remaining by removing the selected channels from selectepackName
        df_remaining = df_myCurrentPack[~df_myCurrentPack['Channel_Name'].isin(df_SelectedChannelsList['Channel_Name'])]
        
        #Get the additional channels from the selected packs that were not part of the current list
        #Lets use the set theory and use DIFFERENCE operation to achieve the results 
        
        #Get the selected channels in a new set from the dataframe of selectec packs
        SelectedChannels=set(df_SelectedChannelsList['Channel_Name'])

        #Get the original channels from current pack list in another set
        OriginalChannels=set(df_myCurrentPack['Channel_Name'])

        #AdditionalChannels = SelectedChannels - OriginalChannels
        AdditionalChannels = SelectedChannels.difference(OriginalChannels)
        #print("mypack count ={}".format(df_myCurrentPack['Channel_Name'].count()))
        #print("addition count={}".format(len(AdditionalChannels)))
        AdditionChannelsCount = len(AdditionalChannels)
        BasePackCount = (df_myCurrentPack['Channel_Name'].count()-1)
        if df_remaining['Channel_Name'].count()==1 :
            if (BasePackCount + AdditionChannelsCount) <=100 :
                RecommendationText = "Add the Network Fee :"
                FeeSub25Chs = 0
            elif (BasePackCount + AdditionChannelsCount) > 100 :
                RecommendationText = "Add the Network Fee + Fee for Subsequent 25 channels:"
                #add the Fee for subsequent 25 channels
                FeeSub25Chs = (df_myCurrentPack['Channel_Name'].count() + len(AdditionalChannels))
                FeeSub25Chs = Calculate_NetworkFee_For25SubsequentChannels(mathObj, FeeSub25Chs)
                dftemp = pd.DataFrame([['Fee for Subsequent 25 Chs',FeeSub25Chs]],
                                      columns=['Channel_Name','Channel_Cost_Per_Month'])
                df_remaining = df_remaining.append(dftemp,ignore_index = True,sort=True)
        else:
            if (BasePackCount + AdditionChannelsCount) <=100 :
                RecommendationText = "Add the Network Fee plus below channels as Alacarte:"
                FeeSub25Chs = 0
            elif (BasePackCount + AdditionChannelsCount) > 100 :
                RecommendationText = "Add the Network Fee + Fee for Subsequent 25 channels plus below channels as Alacarte:"
                #add the Fee for subsequent 25 channels
                FeeSub25Chs = (BasePackCount + AdditionChannelsCount)
                FeeSub25Chs = Calculate_NetworkFee_For25SubsequentChannels(mathObj, FeeSub25Chs)
                dftemp = pd.DataFrame([['Fee for Subsequent 25 Chs',FeeSub25Chs]],
                                      columns=['Channel_Name','Channel_Cost_Per_Month'])
                df_remaining = df_remaining.append(dftemp,ignore_index = True,sort=True)
        
        treeview1.insert('','end',values=("",RecommendationText,"",""))
        
        df_remaining = df_remaining.filter(items=['Channel_Name','Channel_Cost_Per_Month'])
        df_remaining = df_remaining.sort_values(by=['Channel_Cost_Per_Month','Channel_Name'],ascending=[False,True])
        df_remaining = df_remaining.drop_duplicates(keep='first')
        df_remaining = df_remaining.reset_index(drop=True)
        
        blnPrintHeader = False
        for ind in df_remaining.index:
            
            #Display selection in Frame
            if (df_remaining['Channel_Cost_Per_Month'][ind]>0):
                Notes = "Paid" 
            elif (df_remaining['Channel_Name'][ind][:2]=='dd'):
                Notes = "Free=Mandatory"
            else:
                Notes = "Free"
            treeview1.insert('','end',text=str(ind), values=(str(ind+1),
                                                             df_remaining['Channel_Name'][ind],
                                                             str(df_remaining['Channel_Cost_Per_Month'][ind]),
                                                             Notes))
        #compute the cost of alacarte + network + Fee for sub 25 chs  
        CostofSelectedPacks += df_remaining['Channel_Cost_Per_Month'].sum().round(2)
        
        if df_remaining['Channel_Name'].count()==1 :
            if (BasePackCount + AdditionChannelsCount) <=100 :
                RecommendationText = ("Cost of Network Fee (Rs)= {}".format(df_remaining['Channel_Cost_Per_Month'].sum().round(2)))
            elif (BasePackCount + AdditionChannelsCount) >100 :
                BasePackCount -= 1 #Remove the Fee for Subs 25 Chs from the base count 
                RecommendationText = ("Cost of Network Fee + Fee for Subsequent 25 chs (Rs)= \
                {}".format(df_remaining['Channel_Cost_Per_Month'].sum().round(2)))
        else:
            if (BasePackCount + AdditionChannelsCount) <=100 :
                RecommendationText = ("Cost of Network Fee & Alacarte (Rs)= \
                {}".format(df_remaining['Channel_Cost_Per_Month'].sum().round(2)))
            elif (BasePackCount + AdditionChannelsCount) >100 :
                BasePackCount -= 1 #Remove the Fee for Subs 25 Chs from the base count 
                RecommendationText = ("Cost of Network Fee + Fee for subsequent 25 chs & Alacarte (Rs)= \
                {}".format(df_remaining['Channel_Cost_Per_Month'].sum().round(2)))

               
        treeview1.insert('','end',values=("",RecommendationText,df_remaining['Channel_Cost_Per_Month'].sum().round(2),""))    
        #********Next show the individual channels that could not go in with any pack End code*************************
             
        #********Now show the Total cost of new pack and its Savings Begin code*
                
        if df_remaining['Channel_Name'].count() == 1 :
            if (BasePackCount + AdditionChannelsCount) <=100 :
                RecommendationText = ("Total Cost[Pack({})+NWFee({})](Rs)= {}".
                                      format(df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2),
                                             df_remaining['Channel_Cost_Per_Month'].sum().round(2),
                                             CostofSelectedPacks.round(2)))
            elif (BasePackCount + AdditionChannelsCount) >100 :
                RecommendationText = ("Total Cost[Pack({})+(NWFee+FeeSub25chs)({})](Rs)= {}".
                                      format(df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2),
                                             df_remaining['Channel_Cost_Per_Month'].sum().round(2),
                                             CostofSelectedPacks.round(2)))
        else:
            if (BasePackCount + AdditionChannelsCount) <=100 :
                RecommendationText = ("Total Cost[Pack({})+(NWFee+Alac)({})](Rs)= {}".
                                      format(df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2),
                                             df_remaining['Channel_Cost_Per_Month'].sum().round(2),
                                             CostofSelectedPacks.round(2)))
            elif (BasePackCount + AdditionChannelsCount) >100 :
                RecommendationText = ("Total Cost[Pack({})+(NWFee+FeeSub25chs+Alac)({})](Rs)= {}".
                                      format(df_selectedpackNames['Pack_Cost_Per_Month'].sum().round(2),
                                             df_remaining['Channel_Cost_Per_Month'].sum().round(2),
                                             CostofSelectedPacks.round(2)))
        treeview1.insert('','end',values=("",RecommendationText,CostofSelectedPacks.round(2),""))
        
        RecommendationText = ("You Save({}-{}) (Rs)= ".format(CostofMypack.round(2),CostofSelectedPacks.round(2)))
        treeview1.insert('','end',values=("",RecommendationText,(CostofMypack - CostofSelectedPacks).round(2),""))
        
        if len(AdditionalChannels) > 0 :
            RecommendationText ="You had {} channels(includes 25 mandatory) \
                                 in your wishlist but now you will be able to watch {} channel(s)\
                                 additional.".format(BasePackCount,AdditionChannelsCount)
            treeview1.insert('','end',values=("",RecommendationText,"",""))
            RecommendationText ="The Total number of channels in your Recommended pack="
            treeview1.insert('','end',values=("",RecommendationText,int(BasePackCount+AdditionChannelsCount),""))
            RecommendationText ="See the next page for the additional channels list."
            treeview1.insert('','end',values=("",RecommendationText,"",""))
        else:
            RecommendationText ="You had {} channels(includes 25 mandatory) in \
                                 your wishlist which you can retain with the above recommendation.".format(BasePackCount)
            treeview1.insert('','end',values=("",RecommendationText,"",""))
        
        
        RecommendationText = '*All prices above are inclusive of taxes. \
                               Network Fee (NCF) is charged on the basis of the total channel count.'
        treeview1.insert('','end',values=("",RecommendationText,"",""))
        RecommendationText = 'NCF for the first 100 channels is Rs. 153/- per month inclusive of taxes and for every'
        treeview1.insert('','end',values=("",RecommendationText,"",""))
        RecommendationText = 'subsequent 25 channels it is Rs. 23/- per month inclusive of taxes. 1 HD channel is treated as'
        treeview1.insert('','end',values=("",RecommendationText,"",""))
        RecommendationText = '2 SD channels in the channel count.Terms and conditions of the Subscription Contract Apply.'
        treeview1.insert('','end',values=("",RecommendationText,"",""))
         #********Start writing individual channels into FOURTH PAGE End Code*********************************
    #******************Adding the Individual channels(Alacarte Details)******************Page 4 End Code ******************
   
    #******************Adding the Additional channels that were added from the Packs******Page 5 Begin Code ******************
   
        #******Write the additional channels details in the FIFTH PAGE Begin Code**********************************        
        #Show only when there are additional channels to view Begin code
        if len(AdditionalChannels) > 0 :
            #in the last page - Page no 5 show all the additional channels 
            RecommendationText = "************************************************************************"
            treeview1.insert('','end',values=("",RecommendationText,"",""))
            RecommendationText = ("Below are the Additional Channels that you will be able to watch: ")
            treeview1.insert('','end',values=("",RecommendationText,"",""))
            RecommendationText = "************************************************************************"
            treeview1.insert('','end',values=("",RecommendationText,"",""))
            
            RecommendationText = ("Chanel Name")
            treeview1.insert('','end',values=("",RecommendationText,"",""))

            #Iterate through set elements to display the additional channels
            RecommendationText =""
            setIndex = 0
            for channels in AdditionalChannels:
                RecommendationText = channels.title()
                setIndex += 1
                treeview1.insert('','end',values=(str(setIndex),RecommendationText,"",""))
                
    #Show only when there are additional channels to view End code
    #******Write the additional channels details in the FIFTH PAGE End Code*********************************
    #******************Adding the Additional channels that were added from the Packs******Page 5 End Code ******************
       
    #********************Make the recommendations END******************************************************************