import pandas as pd
from letters_creator import Participant, LettersCreator

# Чтобы создать письмо для одного человека
# human = HumanToCongratulate('Учителю', 'Кириенко Денису Павловичу', 'города Москвы', 'школы №179')
# LettersCreator.create_letter(human, 1)
data = pd.read_excel("letters.xlsx")
for i in range(0, len(data)):
    human = Participant('Учителю', data['Subject'].iloc[i], data['Teacher'].iloc[i],
                        data['Location'].iloc[i], data['School'].iloc[i])
    LettersCreator.create_letter(human, i)
# Чтобы создать письма из таблицы
# table = 'https://docs.google.com/spreadsheets/d/1rDlkaDVBmFjElq0BfwaL3Zebo8qeVW3SOPuVOl8tvfo/edit#gid=1516591519'
# LettersCreator.create_letters_from_table(table, 0)
