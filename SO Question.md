I understand that it is possible to `Err.Raise` custom errors in my code to signal problems to the caller. However I'm finding it hard to understand exactly when to do that and I've come across many ways to signal failure to the caller without raising an error:

   1. Return the success/ failure of an operation as a `Boolean` 
       - On its own like the [tryParse pattern][1], but partnered with a seperate error raising method for more info if required
       - With some logging of the exact error, either `Debug.Print` or `MsgBox`, or something similar to `Err.LastDllError` where the caller can query an error state to get more info
       - With no additional info whatsoever

   4. Return an error code / a number representing success or some error
   5. Something else...

All seem to have pros and cons and I'm not sure when to use each and when to raise an error? And if I do raise an error then what information should it convey?

  [1]: https://rubberduckvba.wordpress.com/2019/05/09/pattern-tryparse/