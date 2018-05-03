import curses, time, math, openpyxl

wb = openpyxl.load_workbook("./PycharmProjects/gpl/data_mod.xlsx") # need to be modified depending on the path
#ws = wb.create_sheet(title = "dataSheet")
ws = wb.active
columnTitle = ["Falling Time", "Rising Time", "Falling Speed", "Rising Speed", "Charge(total)", "Number of charge expected", "Charge (experiemental)", "Percent Error"]
for i in range(len(columnTitle)):
    ws.cell(row=1, column=i+1, value=columnTitle[i])

rowIndex = 1
while 1:
    if ws["A{}".format(rowIndex)].value != None:
        rowIndex += 1
    else:
        break

#Curse part

stdscr = curses.initscr()
curses.noecho()
curses.cbreak()

stdscr.keypad(True)
line = 0
timeArr = []

try:
    while True:
        char = stdscr.getch()
        if char == ord('q'):
            timeArr.append(time.time())
            break

        elif char == ord('f'):
            stdscr.addstr(line, 0, str(time.time()))
            timeArr.append(time.time())
            line += 1

finally:
    curses.nocbreak(); stdscr.keypad(False); curses.echo(); curses.endwin()

#Constants

d = 0.00767 #m
rho = 886 #kg/m3
g = 9.80665 #m/s2
eta = 1.867e-5 #poise
b = 8.2e-3
p = 101325 #pa
V = 500 #V
unit_charge = 1.60217662e-19

#Function ###############################

def Charge(v_f, v_r):
    def radius(v_f):
        return math.sqrt((b / (2*p)) ** 2 + (9*eta*v_f)/(2*g*rho)) - b/(2*p)

    def mass(v_f):
        return 4/3 * math.pi * radius(v_f) ** 3 * rho

    return (mass(v_f) * g * (v_f + v_r) * d) / (v_f * V)

def vel(time):
    grid_gap = 0.5e-3
    grid_num = 1
    return grid_gap * grid_num / time

def main():
    global rowIndex
    time_gap = []
    for i in range(len(timeArr)-1):
        time_gap.append(timeArr[i+1] - timeArr[i])
    print("### Real Value: 1.60217662e-19 ###")

    #columnTitle = ["Falling Time", "Rising Time", "Falling Speed", "Rising Speed", "Charge(total)", "Number of charge expected", "Charge (experiemental)", "Percent Error"]

    for i in range(0, int(len(time_gap)/ 2), 2):  # one is up, one is down
        print("Exp #{}: ".format(int(i / 2) + 1), Charge(vel(time_gap[2 * i]), vel(time_gap[2 * i + 2])))
        ws.cell(row=rowIndex, column=1, value=time_gap[2 * i])
        ws.cell(row=rowIndex, column=2, value=time_gap[2 * i + 2])
        ws.cell(row=rowIndex, column=3, value=vel(time_gap[2 * i]))
        ws.cell(row=rowIndex, column=4, value=vel(time_gap[2 * i + 2]))
        charge = Charge(vel(time_gap[2 * i]), vel(time_gap[2 * i + 2]))
        ws.cell(row=rowIndex, column=5, value=charge)
        numCharge = int(charge / unit_charge)
        ws.cell(row=rowIndex, column=6, value= numCharge)
        ws.cell(row=rowIndex, column=7, value= charge / numCharge)
        ws.cell(row=rowIndex, column=8, value= ((charge / numCharge) - unit_charge) * 100 / unit_charge)
        rowIndex += 1

    wb.save("./PycharmProjects/gpl/data_mod.xlsx")

main()
