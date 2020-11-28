import xlwt, xlrd
import os
from pyfmi import load_fmu
from xlutils.copy import copy as xl_copy




# Scheme definition
pro1_mu = -1
pro2_mu = 1
pro3_mu = -1

case='A'

# Common parameters DON'T FORGET TO CHANGE WITH THE SCHEME
# Prosumer 1
pro1_dot_V_sec_in = 10/60 # kg/s, only absolute value
pro1_T_sec_in = 45 + 273.15 # K
pro1_pi = 1


# Prosumer 2
pro2_dot_V_sec_in = 10/60 # kg/s, only absolute value
pro2_T_sec_in = 65 + 273.15 # K
pro2_pi = 1


# Prosumer 3
pro3_dot_V_sec_in = 10/60 # kg/s, only absolute value
pro3_T_sec_in = 45 + 273.15 # K
pro3_pi = 1

# Soil temperature
T_soil = 12 + 273.15 # K


if not os.path.isfile('FMU.xls'):

    print('I couldn\'t find the settings file')

    # Input without configutation file
    # Prosumer 1
    pro1_u_set = [0.8]
    pro1_kappa_set = [0]


    # Prosumer 2
    pro2_u_set = [0]
    pro2_kappa_set = [0.7]


    # Prosumer 3
    pro3_u_set = [0]
    pro3_kappa_set = [0.5]

    case='C'

else:

    book = xlrd.open_workbook('FMU.xls')
    sheet_settings = book.sheet_by_name('Settings')

#    book_copy = xl_copy(book)

    pro1_u_set = sheet_settings.row_values(0)
    pro2_u_set = sheet_settings.row_values(1)
    pro3_u_set = sheet_settings.row_values(2)

    pro1_u_set.pop(0)
    pro2_u_set.pop(0)
    pro3_u_set.pop(0)

    pro1_kappa_set = sheet_settings.row_values(3)
    pro2_kappa_set = sheet_settings.row_values(4)
    pro3_kappa_set = sheet_settings.row_values(5)

    pro1_kappa_set.pop(0)
    pro2_kappa_set.pop(0)
    pro3_kappa_set.pop(0)

book = xlwt.Workbook()
sheet1 = book.add_sheet('FMU')


# bold for header
style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
style.font = font

# header, 26 items
header_values = ['', 'dotQ_1','dotQ_2','dotQ_3','sum_dotQ_loss','T_1_prim_hot',
          'T_1_prim_cold','T_2_prim_hot','T_2_prim_cold','T_3_prim_hot',
          'T_3_prim_cold','T_1_sec_hot','T_1_sec_cold','T_2_sec_hot',
          'T_2_sec_cold','T_3_sec_hot','T_3_sec_cold','dotV_1_prim',
          'dotV_2_prim','dotV_3_prim','dotV_1_sec','dotV_2_sec','dotV_3_sec',
          'Deltap_1_prim','Deltap_2_prim','Deltap_3_prim']

header = sheet1.row(0)

i=0

for column in header_values:
    header.write(i, column, style=style)
    sheet1.col(i).width = 3000
    i=i+1

for i in range(len(pro1_u_set)):
    pro1_u=pro1_u_set[i]
    pro2_u=pro2_u_set[i]
    pro3_u=pro3_u_set[i]

    pro1_kappa=pro1_kappa_set[i]
    pro2_kappa=pro2_kappa_set[i]
    pro3_kappa=pro3_kappa_set[i]


    # Path to the FMU
    model = load_fmu('bdn_2.fmu')

    # Transferring input parameters to the model
    model.set("pro1_dot_V_sec_in", pro1_dot_V_sec_in)
    model.set("pro1_T_sec_in", pro1_T_sec_in)
    model.set("pro1_u", pro1_u)
    model.set("pro1_kappa", pro1_kappa)
    model.set("pro1_mu", pro1_mu)
    model.set("pro1_pi", pro1_pi)
    
    model.set("pro2_dot_V_sec_in", pro2_dot_V_sec_in)
    model.set("pro2_T_sec_in", pro2_T_sec_in)
    model.set("pro2_u", pro2_u)
    model.set("pro2_kappa", pro2_kappa)
    model.set("pro2_mu", pro2_mu)
    model.set("pro2_pi", pro2_pi)

    model.set("pro3_dot_V_sec_in", pro3_dot_V_sec_in)
    model.set("pro3_T_sec_in", pro3_T_sec_in)
    model.set("pro3_u", pro3_u)
    model.set("pro3_kappa", pro3_kappa)
    model.set("pro3_mu", pro3_mu)
    model.set("pro3_pi", pro3_pi)

    # Simulation
    res=model.simulate(final_time=20000)

    # Output of the model
    # Inlet and outlet temperatures at prosumers, deg. C
    PSM1h = round(res["pro1.temPriHot.T"][-1]-273.15,1)
    PSM1c = round(res["pro1.temPriCold.T"][-1]-273.15,1)

    PSM2h = round(res["pro2.temPriHot.T"][-1]-273.15,1)
    PSM2c = round(res["pro2.temPriCold.T"][-1]-273.15,1)

    PSM3h = round(res["pro3.temPriHot.T"][-1]-273.15,1)
    PSM3c = round(res["pro3.temPriCold.T"][-1]-273.15,1)

    # Heat low rate, W
    dotQ_1 = round(res["pro1.plateHEX1.Q1_flow"][-1],1)
    dotQ_2 = round(res["pro2.plateHEX1.Q1_flow"][-1],1)
    dotQ_3 = round(res["pro3.plateHEX1.Q1_flow"][-1],1)

    # Heat losses, W
    Q_loss_1c2c = round(res["pipeCold1.heatPort.Q_flow"][-1],1)
    Q_loss_1h2h = round(res["pipeHot2.heatPort.Q_flow"][-1],1)
    Q_loss_2c3c = round(res["pipeCold2.heatPort.Q_flow"][-1],1)
    Q_loss_2h3h = round(res["pipeHot4.heatPort.Q_flow"][-1],1)
    Q_loss_sum=round(Q_loss_1c2c+Q_loss_1h2h+Q_loss_2c3c+Q_loss_2h3h,1)

    # Secondary side temperature BETTER RENAME IN THE MODEL !!!!!!
    PSM1c_sec = round(res["pro1.temSecHot.T"][-1]-273.15,1)
    PSM1h_sec = round(res["pro1.temSecCold.T"][-1]-273.15,1)
    PSM2c_sec = round(res["pro2.temSecHot.T"][-1]-273.15,1)
    PSM2h_sec = round(res["pro2.temSecCold.T"][-1]-273.15,1)
    PSM3c_sec = round(res["pro3.temSecHot.T"][-1]-273.15,1)
    PSM3h_sec = round(res["pro3.temSecCold.T"][-1]-273.15,1)

    # Volume flow rates (primary side), l/min (density is 1 kg/l)
    dotV_1_pri= round(res["pro1.port_a.m_flow"][-1]*60,1)
    dotV_2_pri= round(res["pro2.port_a.m_flow"][-1]*60,1)
    dotV_3_pri= round(res["pro3.port_a.m_flow"][-1]*60,1)

    # Volume flow rates (secondary side), l/min (density is 1 kg/l)
    dotV_1_sec= round(res["pro1.plateHEX1.m2_flow"][-1]*60,1)
    dotV_2_sec= round(res["pro2.plateHEX1.m2_flow"][-1]*60,1)
    dotV_3_sec= round(res["pro3.plateHEX1.m2_flow"][-1]*60,1)

    # Pressure losses, hPa
    Deltap_1c2c= round(-(res["pro1.port_b.p"][-1]-res["pipeCold1.port_a.p"][-1])/100,1)
    Deltap_1h1c= round(-(res["pipeHot1.port_b.p"][-1]-res["pro1.port_b.p"][-1])/100,1)
    Deltap_1h2h = round(-(res["pipeHotLocal2.port_a.p"][-1]-res["pipeHot2.port_a.p"][-1])/100,1)
    Deltap_2c3c = round(-(res["pro2.port_b.p"][-1]-res["pipeColdLocal2.port_a.p"][-1])/100,1)
    Deltap_2h2c = round(-(res["pipeHot3.port_a.p"][-1]-res["pro2.port_b.p"][-1])/100,1)
    Deltap_2h3h = round(-(res["pipeHot3.port_a.p"][-1]-res["pipeHotLocal5.port_b.p"][-1])/100,1)
    Deltap_3h3c = round(-(res["pipeHot5.port_a.p"][-1]-res["pro3.port_b.p"][-1])/100,1)

    case_designation=case+'-'+str(i+1)
        
    data_values=[case_designation, dotQ_1,dotQ_2,dotQ_3,Q_loss_sum, PSM1h, PSM1c, PSM2h, PSM2c,
             PSM3h, PSM3c, PSM1h_sec, PSM1c_sec, PSM2h_sec, PSM2c_sec,
             PSM3h_sec, PSM3c_sec, dotV_1_pri, dotV_2_pri, dotV_3_pri,
             dotV_1_sec, dotV_2_sec, dotV_3_sec,
             Deltap_1h1c, Deltap_2h2c, Deltap_3h3c] 

    # Starting from the second row
    data = sheet1.row(i+1)

    count=0

    for column in data_values:
        data.write(count, column)
        count=count+1
        
book.save('FMU_output.xls')
