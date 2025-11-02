import win32com.client
import pandas as pd
import datetime
#Constant values
intra_db_link = r'D:\amibroker\AmiBroker intraday\FireAnt_Intr'
eod_db_link =  r'D:\amibroker\AmiBroker\FireAnt_EOD'
class Amibroker():
    def __init__(self):
        '''
        Activate Amibroker application
        '''
        pass
        
    def connect_AB(self):
        AB = win32com.client.Dispatch("Broker.Application")
        AB.Visible =True
        print ('Get Amibroker Application')
        return AB
    
    def active_ticker(self,AB,ticker):

        AD = AB.ActiveDocument
        AD.Name = ticker       
        return
        
    def get_pe_pb(self,AB,ticker):

  
        FnStock = AB.Stocks(ticker)
        
        return round(FnStock.BookValuePerShare,2)  , round(FnStock.EPS,2)  
        

    def Run_analysis(self,AB,program,link_save):
        '''
        '''

        Analysis_program = AB.AnalysisDocs.Open(program)
        Analysis_program.Run(1)
        while Analysis_program.IsBusy:
            pass

        Analysis_program.Export(link_save)
        Analysis_program.Close()
        print ('Run this program: ' + program)
        return 
    
    def load_database(self,AB,link=r'D:\amibroker\AmiBroker\FireAnt_EOD'):

        return AB.LoadDatabase(link)
    def load_layout(self,AB,link=r"C:\Program Files (x86)\AmiBroker\FireAnt_EOD\Layouts\hung_layout.awl"):

        return AB.LoadLayout(link)
    def exit_amibroker(self,AB):
        AB.Quit()
        print ('Exit Amibroker application')
        return

    def check_database_uptodate(self,AB):
        stocks = AB.Stocks
        stock = stocks('MBB')
        last_item = stock.Quotations.count-1
        date = stock.Quotations(last_item).date
        d = date.date()
        today = datetime.date.today()
        if d== today:
            return True
        else:
            return False
        return