from changeworkbook import ChangeWorkbook
from createworkbook import CreateWorkbook


wbname = 'FlowOpt.xlsx'

#create a new, blank workbook
crwb = CreateWorkbook(wbname)
crwb.create()

#change the workbook, fill it
wbname = crwb.get_workbook_name()
chwb = ChangeWorkbook(wbname)

#preparing the workbook for a new cycle
chwb.update_dates('Schedule', 'A3')

#demonstation
chwb.add_to_table('Backlog', 'B', ['Apple', 'Banana', 'Ananas', 'Peas', 'Pods', 'Olives', '1', '2', '3', '4'])
chwb.add_to_table('Backlog', 'D', ['Lady Gaga', 'Tream', 'BonezMC'])

chwb.move_from_backlog_to_schedule('B', ['Apple', 'Banana', 'Olives'])
chwb.move_from_schedule_to_backlog('B', ['Banana'])

#chwb.remove_from_table('Schedule', 4, ['Tream'], 'name')
#chwb.remove_from_table('Backlog', 2, [7, 8], 'index')

chwb.add_to_table('Schedule', 'C', ['Please', 'work', 'cmon', 'and', 'test'])
chwb.remove_from_table('Schedule', 'C', ['cmon'], 'name')

chwb.update_history()
chwb.clear_schedule()

chwb.add_to_table('Schedule', 'A', ['sample1', 'sample2', 'sample3'])
chwb.add_to_table('Schedule', 'E', ['sample4', 'sample5', 'sample6'])
chwb.add_to_table('Schedule', 'G', ['sample7', 'sample8', 'sample9'])

chwb.update_history()
chwb.clear_schedule()

#last step
chwb.savewb(wbname)

# Apples, Bananas, Peas, [3, 4, 6]



#buwb.update_history()          #moves schedule to history
#chwb.update_schedule_dates()   #updates dates in schedule