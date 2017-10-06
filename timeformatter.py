import datetime

def _fix_time(date:str)->str:
    date_split = date.split(' ')
    time_split = date_split[1].split(':')
    hour = int(time_split[0]) % 12
    if hour == 0: hour = 12
    result = str(hour) + ':' +time_split[1] +":" + time_split[2].split('.')[0]
    if int(time_split[0]) > 12:
        result += ' PM'
    else:
        result += ' AM'
    return result

def _fix_date(date:str)->str:
    new_date = date.split(' ')
    new_date = new_date[0].split('-')
    return new_date[1]+'/'+new_date[2]+'/'+new_date[0]



def give_time(date:str)->str:
    return _fix_time(date)

def give_date(date:str)->str:
    return _fix_date(date)


def add_time(date:datetime.datetime, mins: int) -> datetime.datetime:
    return date + datetime.timedelta(minutes = mins)


    







    

    
