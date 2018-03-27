import datetime

def writeToLog(filename, part, err):    
    message = filename + ' - ' + part + ' - ' + str(err)
    final_msg = '{} - {}\n'.format(message, datetime.datetime.now())
    print final_msg
    with open('logFile', 'a') as the_file:        
        the_file.write(final_msg)