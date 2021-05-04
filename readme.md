# Excel VBA undo-history support
__How it works?__ Pretty simple:
- Register user input
- Remember it
- Simulate more user input via __SendKeys__, to perform actions, that would clear undo history, if done from script; do the rest from script

__Problem #1:__ user can interfere with script actions.
__Solution:__ perform a single call to __SendKeys__, providing it with the whole program at once. This way all program actions will be scheduled in advance, and user actions will queue up for them.

__Problem #2:__ how does one perform programmatic actions via __SendKeys__?
__Solution:__ make your program a keyboard-assignable macro, capable of saving its state between calls. Assign it to an unpressable (in a usual way) key combination. __F13 - F15__ keys shall do.

__Problem #3:__ script performs many actions, but they are un/redone one at a time, not in a group.
__Solution:__ use __transactions__. Once a user make some changes, remember them, undo them (via __SendKeys__ of course), simulate entry of transaction index, do what you want, simulate erasure of transaction index. Watch transaction container the rest of time - its value will change as soon as user performs an un/redo. Keep un/redoing until transaction container is empty again.

__Problem #4:__ certain programmatic actions drops user-action history.
__Solution:__
1. Perform write operations on __hidden sheets__, then publish them to user. This approach noticably reduces the list of restricted actions, but, once published, you can't delete or modify those sheets without corrupting undo history, so it has to be postponed until workbook is closed\opened.
2. Use __predefined templates__  to deal with formatting: copy them, populate with data, publish to user. Why not perform formatting from keyboard? Well, I didn't succeed yet.
3. Use __mirroring worksheet functions__, that just reflects data from your hidden sheets.

More details in the code!

## Compatibility
- Originally created for Office 10
- Successful brief test in Office 20

## Shared mode support
__Potentail.__ Supposed to be achievable through a predefined array of templates instead of a single template, copied on demand (shared mode doesn't support copying or deleting sheets at all).

## Known issues
- __Alt+tab__ during execution breaks the whole thing. Guess nothing can be done here.
- __VBA editor is brought to the front on errors,__ and all queued key presses are redirected to it, making debugging a pain. Let your ```On Error Goto```'s be with you.
- __Special key combinations__, like _Shift+Enter_ or _Ctrl+Enter_, provokes additional actions, disrupting the program. Has to be handled manually.
- __Flickering__ while executing. I tried some obvious solutions, but neither worked out, so I just ignore it. Not a really big deal. _Though may turn into a real problem on slow PCs._

## [Demo](https://github.com/danillissimo/Excel-VBA-undo-history-support/blob/main/VBA_undo_history_support_demo.xlsm?raw=true)
Just enter whatever you want in the __ID__ column of the __StaffingBook__ worksheet. It supports:
- IDs from the __Employee__ worksheet
- Empty values
- Unknown values
- Multiple IDs and/or unknown values
- "Vacant" word (case doesn't matter)

An error messasge pops when file is opened - it's a todo, you may need if you gona use it. Just press __skip__ to suppress all of them.

__In case you can't/don't want to launch an untrusted file__ (in which I fully support you), you can perform next actions yourself:
1. Name the workbook module __WB__
2. Create 4 workseets with code names: 
1.__ConstructorFactory__
2.__Employee__
3.__StaffingBook__
4.__TransactionIndexContainer__
3. Put corresponding code inside them
4. Import the rest of modules
5. Fill tables _(all matches are random )_:

__StaffingBook__


| Position         | ID | Name | Sex | Age | Telephone number     |
| ---------------- | -- | ---- | --- | --- | -------------------- |
| Electrician      |    |      |     |     |                      |
| Barrister        |    |      |     |     |                      |
| Housekeeper      |    |      |     |     |                      |
| Speech therapist |    |      |     |     |                      |
| Bookmaker        |    |      |     |     |                      |
| Fashion designer |    |      |     |     |                      |
| Lorry driver     |    |      |     |     |                      |
| Builder          |    |      |     |     |                      |
| Cleric           |    |      |     |     |                      |
| Data processor   |    |      |     |     |                      |
__Employee__

| ID               | Name             | Sex    | Age | Telephone number     |
| ---------------- | ---------------- | ------ | --- | -------------------- |
| #EmptySample#    |                  |        |     |                      |
| #VacantSample#   |                  |        |     |                      |
| #ConflictSample# | Not              | used   | at  | the moment           |
| #ErrorSample#    | INVALID ID       | ###    | ### | ###                  |
| 1                | Harvir Santana   | Male   | 23  | +1 352-224-3574      |
| 2                | Edna Battle      | Female | 54  | +1 210-819-4539      |
| 3                | Codey Mcloughlin | Female | 34  | +1 346-885-6187      |
| 4                | Hashir Horne     | Male   | 35  | +1 405-309-1065      |
| 5                | Aleah Cousins    | Female | 23  | +1 228-865-0047      |
| 6                | Dottie Clifford  | Female | 47  | +1 201-419-9251      |
| 7                | Eduardo Goddard  | Male   | 53  | +1 206-322-8056      |
| 8                | Dev Richards     | Male   | 29  | +1 202-503-8375      |
| 9                | Kadeem Sharp     | Male   | 31  | +1 216-281-8274      |
| 10               | Cherise Pena     | Female | 40  | +1 304-767-3261      |