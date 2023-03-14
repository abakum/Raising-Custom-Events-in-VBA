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
22:58:52 Duelist #1 spun the drum
22:58:52 Duelist #1 heard the sound of a spinning drum
22:58:52 Duelist #2 heard the sound of a spinning drum
22:58:56 Duelist #1 heard the sound of the spinning drum fade away
22:58:56 Duelist #2 heard the sound of the spinning drum fade away

22:58:56 Duelist #1 put the gun to his head
22:58:59 Duelist #1 pulled the trigger
22:58:59 Duelist #1 heard the click of a hammer
22:58:59 Duelist #2 heard the click of a hammer
22:59:01 Duelist #1 handed revolver

22:59:01 Duelist #2 put the gun to his head
22:59:07 Duelist #2 pulled the trigger
22:59:07 Duelist #1 heard the click of a hammer
22:59:07 Duelist #2 heard the click of a hammer
22:59:12 Duelist #2 handed revolver

22:59:12 Duelist #1 put the gun to his head
22:59:17 Duelist #1 pulled the trigger
22:59:17 Duelist #1 heard the click of a hammer
22:59:17 Duelist #2 heard the click of a hammer
22:59:23 Duelist #1 handed revolver

22:59:23 Duelist #2 put the gun to his head
22:59:25 Duelist #2 pulled the trigger
22:59:25 Duelist #1 heard the click of a hammer
22:59:25 Duelist #2 heard the click of a hammer
22:59:28 Duelist #2 handed revolver

22:59:28 Duelist #1 put the gun to his head
22:59:30 Duelist #1 pulled the trigger
22:59:30 Duelist #2 heard a shot
22:59:30 The doctor counts the pulse Duelist #1
22:59:36 The doctor said that Duelist #1 has no pulse
— Finita la comedia! I said to the doctor.
```
