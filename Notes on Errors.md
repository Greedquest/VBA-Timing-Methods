I wanted to be quite rigorous in the way errors are handled. Put simply it comes down to the following ideas:

 - Errors are just another kind or return value of a method. Methods are often used to abstract away implementation details, therefore implementation errors should be condensed to a degree of abstraction which reflects what the caller knows about what goes on inside a method
    - i.e. If I call a method "write to file" I expect abstract errors like "file doesn't exist" rather than specific implementation details "Windows Exception #123 : bytes in wrong place". If the caller has a vague idea what 3 or 4 steps take place in a function, that function should raise 3 or 4 abstract errors for any problems that occur in each of the key steps.
    - Or in the context of my code `CreateTimer` may fail for any one of 100 reasons "DLL File Not Found", `ERROR_NOT_ENOUGH_MEMORY`, `ERROR_PRIVILEGE_NOT_HELD` etc., but the caller only ever sees a `CreateTimerError` and can decide for itself what it wants to do
    - Precise error info should still be logged however as VBA doesn't have an Exception traceback
 - All VBA "errors" can be interpreted as exceptions (here's a [nice explanation][5] of what I understand by Exceptions vs Errors); it is fine to catch all errors and either: 1) condense and rethrow if they are *unchecked exceptions* or 2) use them to alter the course of the program at runtime if they are *checked exceptions*
    - VBA does its best to hide you from actual errors from meddling with pointers and null references, but as far as I'm aware, whenever your code generates something that would be considered an error rather than an exception in other languages, they tend to be *untrappable* (e.g. `Out Of Stack Space error` from too much recursion or 
 - Since VBA doesn't have great control structures for actually handling errors as if they were fully fledged exception objects (i.e no `try...catch`), *checked exceptions* are always caught and enumerated as the method's return value

This is an extension of the `Try-Parse` pattern Matt wrote about. Instead of strictly using `True`/`False`, an error Enum *can* be returned if there are 2 or 3 possible checked exceptions. However as the [MSDN guidelines][6] point out:

> When using this pattern, it is important to define the try
> functionality in strict terms. If the member fails for any reason
> other than the well-defined try, the member must still throw a
> corresponding exception.

So my implementation still re-throws exceptions which don't fall under the `checked` category.


---


Anyway, I'd love to hear some thoughts on either the error handling/raising ethos I'm working towards and/or my implementation of it (am I being consistent?).


I wanted to take a fairly rigorous approach to error handling here. The ethos is:

 - Errors should have same level of granularity as the function which raises them
    - Abstraction is used so that a caller of the TickerAPI doesn't ever need to know about WinAPI internals. It therefore makes no sense to bubble raw dll errors up through the call stack, errors must be of the same degree of abstraction as their function.
    - A better approach therefore is to log precise error details, and then coalesce many errors into a single one for each abstract task a function carries out.
 - It is possible to interpret errors as checked exceptions like in other languages (try-catch style). However because VBA doesn't actually have good control structures for handling errors (like try-catch), the error must be enumerated as a return value of a function, so that proper control structures (Select Case, If..Else, Loops etc.) can be used.
    - This is what the `tryParse` pattern is doing; it says that any of the many possible reasons parsing fails all represent the same error (bad input) and this can be enumerated as a True/False return value, which VBA's control structures can handle a lot better than an raw exception. In languages with try...catch this pattern is essentially equivalent and so has no real benefit, but in VBA it's great.
    - `tryParse` can be expanded to many return values, each representing a given checked exception.
    - These `tryBlah` routines should still bubble unchecked exceptions and errors (the `Else` clause of `try...catch` is often `throw`)
