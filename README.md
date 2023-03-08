# Raising Custom Events in VBA
## Credits
- Mike Wolfe - for [Raising Custom Events in VBA](https://nolongerset.com/raising-custom-events-in-vba/)
- Mikhail Lermontov - for [A Hero of Our Time](https://en.wikipedia.org/wiki/A_Hero_of_Our_Time)

## Usage
- Add to your project module [main.bas](main.bas)
- Add to your project module [пламенныйМотор.bas](пламенныйМотор.bas)
- Add to your project class module [дуэлянт.cls](дуэлянт.cls)
- Add to your project class module [револьверСоднимПатроном.cls](револьверСоднимПатроном.cls)
- Add to your project class module [доктор.cls](доктор.cls)
- Add to your project class module [часы.cls](часы.cls)
- Run КакЭтоБыло()
## Example of Immediate window
```
14:41:57 Дуэлянт №1 крутанул барабан
14:41:57 Дуэлянт №1 услышал звук вращающегося барабана
14:41:57 Дуэлянт №2 услышал звук вращающегося барабана
14:41:59 Дуэлянт №1 услышал, что звук вращающегося барабана затих
14:41:59 Дуэлянт №2 услышал, что звук вращающегося барабана затих

14:41:59 Дуэлянт №1 приставил дуло к своему виску
14:42:04 Дуэлянт №1 нажал на спусковой крючок
14:42:04 Дуэлянт №1 услышал щелчок курка
14:42:04 Дуэлянт №2 услышал щелчок курка
14:42:05 Дуэлянт №1 передал револьвер

14:42:05 Дуэлянт №2 приставил дуло к своему виску
14:42:07 Дуэлянт №2 нажал на спусковой крючок
14:42:07 Дуэлянт №1 услышал выстрел
14:42:07 Доктор слушает пульс у Дуэлянт №2
14:42:10 Доктор сказал, что Дуэлянт №2 убит
– Finita la comedia! – сказал я доктору.
```
