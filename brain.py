import pandas as pd
from ics import Calendar, Event
from dateutil import parser

input = pd.read_excel("harmonogram_example.xlsx", sheet_name="plan")
plan = input.drop(columns=['Forma', 'Dzień tygodnia', 'Numer grupy'])
plan = plan.rename(columns={'Daty zajęć (dd-mm-rr)': 'Daty', 'Budynek i sala': 'Sala'})

for index, _ in plan.iterrows():
    plan.loc[index, 'Sygnatura'] = plan.loc[index, 'Przedmiot'][0:11]
    plan.loc[index, 'Nazwa'] = plan.loc[index, 'Przedmiot'][12:]
    plan.loc[index, 'Prowadzący'] = plan.loc[index, 'Prowadzący'][:-5]
    if pd.isnull(plan.loc[index].Sala):
        plan.loc[index, 'Sala'] = "e-learning"

plan = plan.drop(columns=['Przedmiot'])
plan = plan[plan.columns[[6, 5, 0, 1, 2, 3, 4]]]

export = pd.DataFrame(columns=plan.columns)
for index, row in plan.iterrows():
    daty = plan.loc[index].Daty[:-1].split(';')
    for data in daty:
        plan.loc[index].Daty = data
        export = export.append(plan.loc[index], ignore_index=True)

c = Calendar()

for index, _ in export.iterrows():
    ev = export.loc[index]
    e = Event()
    e.name = ev.Nazwa
    e.begin = parser.parse(''.join((ev.Daty, ' ', ev.Poczatek, ' CEST')), dayfirst=True)
    e.end = parser.parse(''.join((ev.Daty, ' ', ev.Koniec, ' CEST')), dayfirst=True)
    e.location = ev.Sala
    e.description = ev.Prowadzący
    c.events.add(e)

file = open('plan.ics', 'w')
file.writelines(c)
file.close()
