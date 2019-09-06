As anyone who's ever worked for any significant length of time with [WinAPI Timers][1] probably knows, they can be some of the most temperamental, unstable and crash-inducing elements of a project, and as such are very tricky to work with. Instances of random, unstoppable printing to the debug window, frequent screen freezes or outright crashes are probably commonplace.

 The purpose of this project is to understand what causes these problems and to avoid them happening. Here I present the code which I've come up with to try and mitigate the issues and make working with the APIs a bit more straightforward.

First though, some background*...

<sup>* You could skip/skim this if you're familiar with how message queues and especially `WM_TIMER` messages work. Also quick disclaimer; I've only freshly learnt this stuff, I've done my best to make sure it's correct, but do correct any errors/ add clarifications as you see fit. </sup>

#Of Message Queues and Timers

Windows Applications (like Excel or whatever other application is hosting your VBA code) use *threads* to host the different tasks they want to run. One way in which applications can communicate with each other (or themselves) is by sending instructions (messages) to a thread. Each thread gets a unique message queue associated with it, and when a message is sent to a thread it gets added to the queue. 

The thread's [message loop][2] slowly works through this queue of instructions, generally just forwarding them on to different windows on that thread using the `DispatchMessage` function according to the `hWnd` parameter of the message  - these messages get handled by the destination window's [`WNDPROC`][3], a function that can be subclassed (overwritten) to respond differently to the messages sent. However not all messages are handled by an `WNDPROC`, some messages include a pointer to a custom function that can handle the message...

VBA (in)famously gets executed from the host application's Main/UI thread. From the [VBA tag wik][4]i:

> Most VBA hosts run VBA code on the main/UI thread, so it's perfectly
> normal that a long-running VBA macro makes the host application's main
> window caption say "(not responding)". In that state, the VBA code is
> running, and the host application is no longer handling Windows
> messages

Messages from timers must also go through this thread if they want to reach VBA code.

---

There are two varieties of timer exposed by the Windows API; thread based and message based timers. 

 - Thread based timers are the bees knees; the idea is that you create another thread with say the [`CreateThread` function][5]. You can then tell this thread to sleep for a while, and then when it wakes up to call a method in your main thread. This way your main thread can be busy doing other stuff during the wait (rather than a blocking sleep). However I'm yet to find anyone who's had much success with this, and besides it's probably overkill for what I imagine most people want to do with timers in VBA
 - Message based timers meanwhile generate `WM_TIMER` messages in a thread's message queue at set intervals, and these are processed asynchronously. Someone's still doing the waiting somewhere, but this is handled internally by Windows so there's no need to worry about it.

`SetTimer` can be used to make a message based timer. There is quite a lot of flexibility in the way you do this, but the two most common I see are

    timerID = SetTimer(0,0,delayMillis,AddressOf myTimerProc)

    timerID = SetTimer(Application.Hwnd,ObjPtr(keyOrArgs),delayMillis,AddressOf myTimerProc)

Both declarations will create Timers that live on the thread which made them (which for VBA code will be a thread provided by the host application). This means that even if you forget to call `KillTimer`, or Excel crashes, Windows will kill those timers for you eventually (when it releases the host applications memory), so no need to worry about leaving them floating about by accident.

Both declarations also have a `TIMERPROC` set. This means that when the thread's message loop gets to the `DispatchMessage` function, rather than forwarding to a window handle's `WNDPROC`, the message is popped from the queue and sent to the `TIMERPROC`.

The difference between the declarations is that for the first one, Windows will generate a unique ID (unique to the thread at least) for you. For the second approach, you specify the ID and it gets stored in a list which is associated with that window handle. This allows you to specify whatever ID you want, and often a pointer to an object is used because

 - It is guaranteed to be unique (certainly within the scope of the parent handle)
 - It can be dereferenced by the TimerProc (so you can pass data around)



### Of `WM_TIMER` messages

`WM_TIMER` messages have some quirks, being aware of these helps explain some of the unpredictable behaviour of timers:

 - They allow an optional callback to be specified (the TimerProc).

- This TimerProc is called directly by Windows; Windows does not know how to deal with VBA Errors, and this fact is probably the number one cause of crashes - TimerProcs that raise errors which can't be handled.
- Another problem is passing invalid function pointers to `SetTimer`; when these are dereferenced Windows probably sends an exception to Excel which it does not know how to handle - causing a crash.
- A third issue associated with the TimerProc is that the Object used to generate the unique ID may fall out of scope/ be destroyed before the callback is invoked; if VBA then dereferences the pointer we'll get a crash.

 - Unlike other messages, WM_TIMER messages are generated on the fly. When a Timer expires (its period elapses), rather than creating a message and posting it to the queue, it instead sets a flag on the queue which means "when there are no more messages to process, create a WM_TIMER message". Crucially this means that messages don't (shouldn't) build up ([sometimes they do][6], it depends on exactly what the message loop does - but not in my testing).

- I mention this because I've heard it said that the reason Excel crashes when you have a timer open and edit a cell is because "Too many timer messages build up". This is not the case, the real reason is because [the Excel Object Model is not designed to work with asynchronous code][7] (like timer callbacks); i.e the timer callback is trying to write to a cell at the same time as Excel is handling user input (and that raises an unhandled error)

 - Killing a timer will prevent any new messages being made; it will clear the flag that is set to generate more. However it will not remove messages already in the queue (which got there by `Peek`ing when a flag is set, or `Post`ing manually - the queue is a public place so it makes sense to be aware of this). If a TimerProc *is not* set then this means WM_TIMER messages will continue to be `Dispatch`ed to their destination WindowProcs. However if a TimerProc is set, Windows seems to be clever enough [not to invoke the TimerProcs][8] of messages which were generated by timers that have since been killed. This makes messages with a TimerProc a little more predictable than those without.

---

# TL;DR (also if you did read it)

With a better understanding of what's going on behind the scenes, it's possible to do a Q. & A. of common problems and the solutions which this code uses.

 1. 


  [1]: https://docs.microsoft.com/en-gb/windows/win32/winmsg/timers
  [2]: https://docs.microsoft.com/en-gb/windows/win32/winmsg/using-messages-and-message-queues#creating-a-message-loop
  [3]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms633573(v%3Dvs.85)
  [4]: https://codereview.stackexchange.com/tags/vba/info
  [5]: https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createthread
  [6]: https://devblogs.microsoft.com/oldnewthing/20160624-00/?p=93745
  [7]: https://support.microsoft.com/en-gb/help/2800327/limitation-of-asynchronous-programming-to-the-excel-object-model
  [8]: https://stackoverflow.com/q/57134016/6609896