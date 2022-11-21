#!/usr/bin/env python
""" Communication between E3 and Python with E3DataAccess
This program is using several tags to demonstrate the ability to handle several
between E3 and Python with E3DataAccess. It should run simulatenously with the
Elipse E3 domain of the same name.
This program is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software
Foundation, either version 3 of the License, or (at your option) any later
version.
This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.
GNU General Public License : <http://www.gnu.org/licenses/>.
"""

__author__ = "Julien Goy"
__contact__ = "julien@elipse.com.br"
__date__ = "2022/11/07"
__status__ = "Prototype"
__version__ = "0.0.3"

# import libraries
from win32com.client import Dispatch, DispatchWithEvents
from pythoncom import PumpWaitingMessages
from sys import exit, argv
from datetime import datetime

#class Data
class Data:
    tagread = ""
    tagwrite = ""

    def __init__(self, i=1, counter=0):
        self.tagread = 'Data.XoTag' + str(i) + '.Data.Value'
        self.tagwrite = 'Data.XoTag' + str(i) + '.Counter.Value'
        self.counter = counter
        self.quality = 192

    def register(self, eComCall):
        register = eComCall.RegisterCallback(self.tagread)
        print("Registered tag : {} : {}".format(self.tagread, register))
        return register

    def increment(self, eComCall):
        self.counter +=1
        result = eComCall.WriteValue(self.tagwrite, datetime.now(), self.quality, self.counter)
        return result

# define Application Events
class ApplicationEvents:
    eComCall2 = None
    tags = None

    # define an event to fire when value change on tags
    def OnValueChanged(obj, path, timestamp, quality, value):

        for tag in tags:
            if tag.tagread == path:
                result = tag.increment(eComCall2)
                print('{} : {} : {} : {}'.format(tag.tagwrite, tag.counter, value, result))

if __name__ == '__main__':

    for i in range(1, len(argv)):
        print('argument:', i, 'value:', argv[i])

    # Get the active instance of E3 Data Access
    e3dataaccess = 'E3DataAccess.E3DataAccessManager.1'

    # Create a Dispatch based COM object for read/write operation
    eComCall2 = Dispatch(e3dataaccess)

    # Create a Dispach based COM object that to fire events to our defined class
    eComCall = DispatchWithEvents(e3dataaccess, ApplicationEvents)

    # Create 20 tags
    tags = []
    for i in range(20):
        tags.append('tag'+str(i))
        tags[i] = Data(i+1)
        tags[i].register(eComCall)

    ApplicationEvents.eComCall2 = eComCall2
    ApplicationEvents.tags = tags

    # define initalizer
    keepOpen = True

    # look to receive events
    while keepOpen:

        # pumps all waiting messages for the current thread
        PumpWaitingMessages()

        try:
            keepOpen = True

        except:
            # if there is an error close and exit the script
            keepOpen = False
            eComCall = None
            eComCall2 = None
            exit()