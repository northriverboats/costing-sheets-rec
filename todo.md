# ToDo's
* cli output for use with frontend
    * use tabs for status
    * update buffer for each section
    * ouptut all sections on every refresh

```
status = None

class Status():
    def __init__(self):
        self.__percent = 0
        self.__file = ''
        self.__section = ''

    def __str__(self):
        return "{}\t{}\t{}".format(
            self.__percent, 
            self.__file,
            self.__section)

    @property
    def percent(self):
        return self.__percent

    @percent.setter
    def percent(self, value):
        self.__percent = value

    @property
    def file(self):
        return self.__file

    @file.setter
    def file(self, value):
        self.__file = value

    @property
    def section(self):
        return self.__section

    @section.setter
    def section(self, value):
        self.__section = value

status = new Status()
status.percent = 25
status.file  = r"K:\Links\2020\Boats\boats.pickle"
status.section = 'OUTFITTING'
print(status)

```