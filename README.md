# Raising Custom Events in VBA
## Credits
- Mike Wolfe - for [Raising Custom Events in VBA](https://nolongerset.com/raising-custom-events-in-vba/)
- Mikhail Lermontov - for [A Hero of Our Time](https://en.wikipedia.org/wiki/A_Hero_of_Our_Time)

## Usage
- Add to your project module [main.bas](main.bas)
- Add to your project module [heart.bas](heart.bas)
- Add to your project class module [revolverWithSingleCartridge.cls](revolverWithSingleCartridge.cls)
- Add to your project class module [duelist.cls](duelist.cls)
- Add to your project class module [doctor.cls](doctor.cls)
- Run howItWas() from `main`
## Example of Immediate window
```
15:15:30 Duelist #1 spun the drum
15:15:30 Duelist #1 heard the sound of a spinning drum
15:15:30 Duelist #2 heard the sound of a spinning drum
15:15:33 Duelist #1 heard the sound of the spinning drum fade away
15:15:33 Duelist #2 heard the sound of the spinning drum fade away

15:15:33 Duelist #2 put the gun to his head
15:15:35 Duelist #2 pulled the trigger
15:15:35 Duelist #1 heard the click of a hammer
15:15:35 Duelist #2 heard the click of a hammer
15:15:40 Duelist #2 handed revolver

15:15:40 Duelist #1 put the gun to his head
15:15:45 Duelist #1 pulled the trigger
15:15:45 Duelist #1 heard the click of a hammer
15:15:45 Duelist #2 heard the click of a hammer
15:15:47 Duelist #1 handed revolver

15:15:47 Duelist #2 put the gun to his head
15:15:53 Duelist #2 pulled the trigger
15:15:53 Duelist #1 heard the click of a hammer
15:15:53 Duelist #2 heard the click of a hammer
15:15:59 Duelist #2 handed revolver

15:15:59 Duelist #1 put the gun to his head
15:16:04 Duelist #1 pulled the trigger
15:16:04 Duelist #1 heard the click of a hammer
15:16:04 Duelist #2 heard the click of a hammer
15:16:08 Duelist #1 handed revolver

15:16:08 Duelist #2 put the gun to his head
15:16:12 Duelist #2 pulled the trigger
15:16:12 Duelist #1 heard a shot
15:16:12 The doctor counts the pulse Duelist #2
15:16:18 The doctor said that Duelist #2 has no pulse
— Finita la comedia! I said to the doctor.
```
