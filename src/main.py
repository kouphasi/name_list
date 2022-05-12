import openpyxl as opxl
from operator import itemgetter
from action.addfreshman import addFresh
from action.change_debate_to_discussion import change_DebToDis
from action.divide import devide_3
from action.erase_nosub import erase_noSub
from action.renew import add_under4
from action.sort import make_membersheet
from action.auto_width import doAuto

newBook = opxl.Workbook()
newBook.save(r"goal\2022_ESS名簿.xlsx")

addFresh()
add_under4()
erase_noSub()
change_DebToDis()
make_membersheet()
devide_3()
doAuto()