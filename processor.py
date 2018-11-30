# this script acts like a big switch statement
# it takes the organization name
#send the file path, month, year, and organization name to hardi
#to convert the file
from math import ceil
import hardi
import sys

def process(orgName,year,month,file_names,save_file_path):

    stat_message=''
    result_message=''
    open_file_path=''

    if len(file_names)==1:
        open_file_path=file_names[0]

    if orgName == '2J Supply' or orgName == 'Illco':
        df, row = hardi.twoJsupply(open_file_path, orgName, year, month)

    elif orgName == 'ABR Wholesalers' or orgName == 'APR Supply' or orgName=='CFM Equipment Distributors':
        df, row = hardi.ABRWholeSalers(open_file_path, orgName, year, month)

    elif orgName=='AC Pro':
        df, row=hardi.acpro(open_file_path,orgName,year,month)

    elif orgName=='ACR Supply Company':
        df,row=hardi.acrsupply(open_file_path,orgName,year,month)

    elif orgName == 'American Air Distributing':
        df, row = hardi.AmericanAirDistributing(open_file_path, orgName, year, month)

    elif orgName=='Aireco Supply':
        df,row=hardi.aireco(open_file_path,orgName,year,month)

    elif orgName == 'API of NH':
        df, row = hardi.APIofNH(open_file_path, orgName, year, month)

    elif orgName=='Associated Equipment Company':
        df,row=hardi.associatedequipment(open_file_path,orgName,year,month)

    elif orgName == 'Auer Steel':
        df, row = hardi.AuerSteel(open_file_path, orgName, year, month)

    elif orgName == 'Best Choice' or orgName == 'Dunphey and Associates Supply Co.':
        df, row = hardi.BestChoice(open_file_path, orgName, year, month)

    elif orgName=='Behler-Young':
        df,row=hardi.behleryoung(open_file_path,orgName,year,month)

    elif orgName=='Benoist Brothers':
        df,row=hardi.benoist(open_file_path,orgName,year,month)

    elif orgName == 'Capco Energy Supply':
        df, row = hardi.capco(open_file_path, orgName, year, month)

    elif orgName == 'Carr Supply' or orgName=='Weathertech Distributing Co':
        df, row = hardi.CarrSupply(open_file_path, orgName, year, month)

    elif orgName == 'Century AC Supply':
        df, row = hardi.CenturyACSupply(open_file_path, orgName, year, month)

    elif orgName=='cfm Distributors Inc':
        df,row=hardi.cfmdis(open_file_path,orgName,year,month)

    elif orgName == 'Charles D Jones Company':
        df, row = hardi.cdj(file_names, orgName, year, month)

    elif orgName == 'Comfort Air Distributing':
        df, row = hardi.CAD(open_file_path, orgName, year, month)

    elif orgName=='Comfort Supply':
        df,row=hardi.comfortsupply(open_file_path,orgName,year,month)

    elif orgName == 'Corken Steel Products Company':
        df, row = hardi.CSPC(open_file_path, orgName, year, month)

    elif orgName=='Crescent Parts' or orgName=='Key Refrigeration Supply':
        df,row=hardi.crescent(open_file_path,orgName,year,month)

    elif orgName == 'Design Air':
        df, row = hardi.designAir(open_file_path, orgName, year, month)

    elif orgName == 'Distributor Corporation of New England':
        df, row = hardi.dcne(open_file_path, orgName, year, month)

    elif orgName == 'Duncan Supply Company':
        df, row = hardi.dsc(file_names, orgName, year, month)

    elif orgName == 'Ed\'s Supply':
        df, row = hardi.edSupply(open_file_path, orgName, year, month)

    elif orgName == 'Ferguson':
        df, row = hardi.ferguson(open_file_path, orgName, year, month)

    elif orgName=='G.W. Berkheimer Company':
        df,row=hardi.gwberkheimer(open_file_path,orgName,year,month)

    elif orgName=='Geary Pacific Supply':
        df,row=hardi.gearypacific(open_file_path,orgName,year,month)

    elif orgName=='The Granite Group':
        df,row=hardi.granite(open_file_path,orgName,year,month)

    elif orgName=='Gustave A Larson Company':
        df,row=hardi.gustave(open_file_path,orgName,year,month)

    elif orgName == 'Heating and Cooling Supply Co. Inc.':
        df, row = hardi.hcs(open_file_path, orgName, year, month)

    elif orgName == 'Hercules Industries':
        df, row = hardi.HerculesIndustries(open_file_path, orgName, year, month)

    elif orgName == 'HVAC Distributors':
        df, row = hardi.HVACDistributors(open_file_path, orgName, year, month)

    elif orgName == 'HVAC Sales & Supply Company':
        df, row = hardi.HSSC(open_file_path, orgName, year, month)

    elif orgName=='Interline Brands':
        df,row=hardi.interline(open_file_path,orgName,year,month)

    elif orgName == 'Johnson Supply':
        df, row = hardi.johnsonSupply(open_file_path, orgName, year, month)

    elif orgName == 'Johnstone Supply - Popma':
        df, row = hardi.jsPopma(open_file_path, orgName, year, month)

    elif orgName == 'Koch Air':
        df, row = hardi.kochAir(open_file_path, orgName, year, month)

    elif orgName == 'Locke Supply':
        df, row = hardi.lockeSupply(open_file_path, orgName, year, month)

    elif orgName == 'Lohmiller & Company':
        df, row = hardi.lohmiller(open_file_path, orgName, year, month)

    elif orgName=='Luce Schwab & Kase':
        df,row=hardi.luce(open_file_path,orgName,year,month)

    elif orgName=='M&A Supply':
        df,row=hardi.m_and_a(open_file_path,orgName,year,month)

    elif orgName=='McCall\'s Supply Company':
        df,row=hardi.mccall(open_file_path,orgName,year,month)

    elif orgName=='Meier Supply':
        df,row=hardi.meier(open_file_path,orgName,year,month)

    elif orgName == 'Minnesota Air':
        df, row = hardi.minnesota(open_file_path, orgName, year, month)

    elif orgName=='Morsco':
        df,row=hardi.morsco(open_file_path,orgName,year,month)

    elif orgName=='Mingledorff\'s':
        df,row=hardi.mingledorff(open_file_path,orgName,year,month)

    elif orgName=='Munch Supply Company':
        df,row=hardi.munch(open_file_path,orgName,year,month)

    elif orgName=='NB Handy':
        df,row=hardi.nbhandy(open_file_path,orgName,year,month)

    elif orgName=='O\'Connor Company':
        df,row=hardi.oconnor(open_file_path,orgName,year,month)

    elif orgName=='Peirce Phelps':
        df,row=hardi.peirce(open_file_path,orgName,year,month)

    elif orgName == 'Shearer Supply':
        df, row = hardi.shearer(open_file_path, orgName, year, month)

    elif orgName=='Sid Harvey\'s':
        df,row=hardi.sidharvey(open_file_path,orgName,year,month)

    elif orgName=='Sigler Wholesale Distributors':
        df,row=hardi.sigler(open_file_path,orgName,year,month)

    elif orgName=='South Central Company':
        df,row=hardi.southcentral(open_file_path,orgName,year,month)

    elif orgName == 'Standard Supply':
        df, row = hardi.standard(open_file_path, orgName, year, month)

    elif orgName == 'Star Supply Company':
        df, row = hardi.starSupply(open_file_path, orgName, year, month)

    elif orgName=='RE Michel':
        df,row=hardi.remichel(open_file_path,orgName,year,month)

    elif orgName=='Refrigeration Sales Corp':
        df,row=hardi.refrigeration(open_file_path,orgName,year,month)

    elif orgName=='Robert Madden':
        df,row=hardi.robertmadden(open_file_path,orgName,year,month)

    elif orgName=='Robertson Heating Supply':
        df,row=hardi.robertson(open_file_path,orgName,year,month)

    elif orgName=='Team Air Distributing':
        df,row=hardi.teamair(open_file_path,orgName,year,month)

    elif orgName == 'Temperature Systems':
        df, row = hardi.temperature(open_file_path, orgName, year, month)

    elif orgName=='Thos. Somerville Company':
        df,row=hardi.thossomerville(open_file_path, orgName, year, month)

    elif orgName=='Total Air Supply':
        df,row=hardi.totalAir(open_file_path, orgName, year, month)

    elif orgName == 'Williams Distributing':
        df, row = hardi.williams(open_file_path, orgName, year, month)

    elif orgName == 'Winsupply':
        df, row = hardi.winsupply(open_file_path, orgName, year, month)

    elif orgName == 'Airefco':
        df, row = hardi.airefco(open_file_path, orgName, year, month)

    elif orgName == 'Capitol Group':
        df, row = hardi.capitol(open_file_path, orgName, year, month)

    else:
        result_message='ohps! Lord Varis hasn\'t work on it yet'
        return stat_message,result_message

    #df means dataframe
    #my script heavily depends on Python's Pandas library
    # here df is in converted version, i just simply write it to a csv file
    # save the file
    df.to_csv(save_file_path, index=False,encoding='utf-8')

    # summary
    # do some simple statistic test
    count = df['Total_Unit_Sales'].count()
    sum = df['Total_Unit_Sales'].sum()

    stat_message='Submitted data: COUNT=%d, SUM=%d' % (count, sum)

    quarter = ceil(month / 3)

    #stat check to use for quarterly submissions 
    if orgName=='Capitol Group' or orgName=='Airefco' or orgName=='Meier Supply':
        if df['Quarter'].sum() != quarter * row:
            result_message="There is data from different quarters"
        else:
            result_message='Good to go!'
    else:
        if df['Month'].sum()!=month*row:
            result_message="There is data from different month"
        else:
            result_message='Good to go!'

    return stat_message,result_message
