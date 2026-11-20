import random
import numpy as np
from scipy import stats
import win32com.client
import pythoncom
import os

def cpk_estudio(upl, lwl):
    pythoncom.CoInitialize()
    try:
        return _cpk_estudio(upl, lwl)
    finally:
        pythoncom.CoUninitialize()

def _cpk_estudio(upl, lwl):
    p_n = 0.0
    out_x = 1
    out_r = 1
    j = 0

    u_lmt = upl
    l_lmt = lwl
    mean = (u_lmt + l_lmt) / 2

    while (((out_x > 0) or (out_r > 0)) and (j < 3)):
        number = random.uniform(2.5, 7.5)
        dev_std = (u_lmt - mean) / (3 * number)
        print(f'CPK generado: {number}')

        while p_n < 0.05:
            gauss = []

            for i in range(125):
                gauss.append(random.gauss(mean, dev_std))

            stat_n, p_n = stats.shapiro(gauss)
            #dev_std_calculated = np.std(gauss)
            print(f'Normal distribution : {p_n}')

        mtb = win32com.client.Dispatch("Mtb.Application.1")
        mtb.UserInterface.Visible = False
        mtb.UserInterface.DisplayAlerts = False

        columns = mtb.ActiveProject.ActiveWorksheet.Columns
        column = columns.Add()
        column.Name = "CPK"
        column.SetData(gauss)

        columns.Add().Name = "Mean1" 
        columns.Add().Name = "StDev1"
        columns.Add().Name = "PPL"
        columns.Add().Name = "PPU"
        columns.Add().Name = "Subgroup"
        columns.Add().Name = "Mean2"
        columns.Add().Name = "Numerator"
        columns.Add().Name = "Denominator"
        columns.Add().Name = "Sp"
        columns.Add().Name = "Constant4"
        columns.Add().Name = "SdWithin"
        columns.Add().Name = "Mean3"
        columns.Add().Name = "CPL"
        columns.Add().Name = "CPU"
        columns.Add().Name = "Range1"
        columns.Add().Name = "ByVar4"
        columns.Add().Name = "Mean4"
        columns.Add().Name = "Mean5"
        columns.Add().Name = "Mean6"
        columns.Add().Name = "Mean7"
        columns.Add().Name = "D4"
        columns.Add().Name = "UCL_R"
        columns.Add().Name = "LCL_R"
        columns.Add().Name = "A2"
        columns.Add().Name = "UCL_X"
        columns.Add().Name = "LCL_X"
        columns.Add().Name = "Status_X"
        columns.Add().Name = "Status_R"
        columns.Add().Name = "Condition_X"
        columns.Add().Name = "Condition_R"

        command = f"Sixpack 'CPK' 5; Lspec {l_lmt}; Uspec {u_lmt}; Pooled; AMR; CCRbar; CCSbar; CCAMR; UnBiased; OBiased; Breakout 25; Target {mean}; Toler 6; CStat; Test 1."
        pp = f"Statistics 'CPK'; Mean 'Mean1'; StDeviation 'StDev1'. Let 'PPL' = ('Mean1'- {l_lmt})/((6/2)*'StDev1') \n Let 'PPU' = ({u_lmt}-'Mean1')/((6/2)*'StDev1')"
        columns.Item("Constant4").SetData(0.997503)
        cp = f"Set 'Subgroup' \n 1( 1 : 25 / 1 )5 \n End. Statistics 'CPK'; By 'Subgroup'; Expand; Mean 'Mean2'. Formula 'Numerator' = SUM(('CPK'-'Mean2')^2) \n Formula 'Denominator' = 25*(5-1) \n Formula 'Sp' = SQRT('Numerator'/'Denominator') \n Formula 'SdWithin' = 'Sp'/'Constant4' \n Statistics 'CPK'; Mean 'Mean3'. Formula 'CPL' = ('Mean3'-{l_lmt})/((6/2)*'SdWithin') \n Formula 'CPU' = ({u_lmt}-'Mean3')/((6/2)*'SdWithin')"
        columns.Item("D4").SetData(2.115)
        columns.Item("LCL_R").SetData(0)
        columns.Item("A2").SetData(0.577)
        charts = f"Statistics 'CPK'; \n By 'Subgroup'; \n Expand; \n Range 'Range1'. Statistics 'Mean2' 'Range1'; By 'Subgroup'; GValues 'ByVar4'; Mean 'Mean4'-'Mean5'. Statistics 'Mean4' 'Mean5'; Mean 'Mean6'-'Mean7'. Let 'UCL_R' = 'Mean7' * 'D4'\n Let 'UCL_X' = 'Mean6' + 'A2' * 'Mean7' \n Let 'LCL_X' = 'Mean6' - 'A2' * 'Mean7' \n Formula 'Status_X' = ('Mean4'  > 'UCL_X') Or ('Mean4' < 'LCL_X') \n Formula 'Status_R' = ('Mean5'  > 'UCL_R') Or ('Mean5' < 'LCL_R') \n Formula 'Condition_X' = SUM('Status_X') \n Formula 'Condition_R' = SUM('Status_R')"

        mtb.ActiveProject.ExecuteCommand(command)
        mtb.ActiveProject.ExecuteCommand(pp)
        mtb.ActiveProject.ExecuteCommand(cp)
        mtb.ActiveProject.ExecuteCommand(charts)

        output = mtb.ActiveProject.Commands.Item(1).Outputs.Item(1).Graph
        cur_dir = os.path.dirname(os.path.abspath(__file__))
        save_path = os.path.join(cur_dir, "cpkRandomGaus.png")
        output.SaveAs(save_path)

        ppl = format(float(columns.Item("PPL").GetData(1, 1)), ".2f")
        ppu = format(float(columns.Item("PPU").GetData(1, 1)), ".2f")
        if ppl < ppu:
            PPK = ppl
        else:
            PPK = ppu
        print("PPK:", PPK)
        cpl = format(float(columns.Item("CPL").GetData(1, 1)), ".2f")
        cpu = format(float(columns.Item("CPU").GetData(1, 1)), ".2f")
        if cpl < cpu:
            CPK = cpl
        else:
            CPK = cpu
        print("CPK:", CPK)

        out_x = int(columns.Item("Condition_X").GetData(1, 1))
        out_r = int(columns.Item("Condition_R").GetData(1, 1))
        print(f"Datos fuera de rango X Chart: {out_x} \nDatos fuera de rango R Chart: {out_r}")
        mtb.Quit()
        p_n = 0.0
        j+=1

    if (out_r > 0 and out_x):
        return(f"Estudio erroneo, {out_r} subgrupos fuera de limites en R Chart y {out_x} subgrupos fuera de limite en X Chart")
    elif (out_r > 0 ):
        return(f"Estudio erroneo, {out_r} subgrupos fuera de limites en R Chart")
    elif (out_x > 0 ):
        return(f"Estudio erroneo, {out_x} subgrupos fuera de limites en X Chart")
    else:
        return("Estudio finalizado satisfactoriamente")

#cpk_estudio(0.312, 0.125)