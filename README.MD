
# BinTri
 -------
### A fast binary search program which exploits additional return values.
### The change means the number of compares the function needs practically halves.
### It could be considered as using **trinary logic** rather than binary.

                           3 STATES   i.e < 0, 0 , > 0

           Return values such as this are common in C++ but less so in other
           languages.

           For binary searches, it is ideal at identifying a midpoint value between
           high and low values. In maths, it is also called balanced ternary and
           was used in some early calculating machines. Whereas binary deals in 
		   bits, trinary deals in trits of information.

           The function is particularly suited to string searching and any lists
           where the compare against a target value requires a large amount of
           processing time.

           It is written in VBA (Visual Basic for Applications) which is built into 
		   Microsoft Excel. It's intended to be used as a viable alternative to needing
		   an external database and as a proof of concept in the use of trinary logic.

           BinTri is just the standard Binary Search found on Wikipedia which has
           been adapted to the 3 state compare values.

           On average this change makes BinTri 66% faster at Case Insensitive
           string searches.


           The code is minimal and easily readable. ...Richard Marsden 2021



 NOTES
 -----

 1) The 3 state comparison is implemented with.....
   
     INTEGERS: (on 32 bit positive signed) A test subtraction.
    
     ret = num1 - num2 
	  
	 ret is negative, 0 or positive result  (3 State)
and   
   
      STRINGS: (Unicode UTF16)
    
      For strings, the VBA function StrComp is used for comparison.
              
			  ret = StrComp(string1, string2)
			  
      The function returns        -1  where string1 is less    than string2
                                   0  where string1 is equal   to   string2
                                   1  where string1 is greater than string2

      NOTE: C++ has a similar function Strcmp (ASCII) and Wcscmp (Unicode).

      Return values in C++ are: a negative value, 0, a positive value (3 State)



-----------------


1) Three binary search programs are listed in this document although seven were tested for speed.

The main candidates to judge BinTri (abbr. **TRI**) against is the standard binary search named BinFind 
 (**STANDARD**) and MonoBlock (**MONO**) which is a slight variation of the alternate binary search 
mentioned in Wikipedia.


2) For string searches TRI has about 66% speed advantage over STANDARD binary search and a
   9% advantage over MONO if case insensitive option is ON (Option Compare Text).

   With case sensitive compares (Option Compare Binary) this reduces to about a 41% speed
   advantage over STANDARD and 7% over MONO.

3) No real change in this advantage ratio with either large (10million) or small arrays (2500)

4) With 10% of searches set to "No Find" the overall speed loss is not significant. 
  
   STANDARD and TRI do however lose their main advantage over MONO where there isn't a match 
   ("No Find") as they are unable to match a value early.
   
   TRI will proceed with all possible matches before a "No Find" result. The compare count in 
   this case will be **one less** than MONO. This is because MONO uses a final equality compare on 
   exiting its loop.
   

5) No real change with any of them if 10% of searches have a Unicode lower case mismatch
   introduced. Note 4) applies.

6) No real speed difference if starve computer of memory with a huge dummy array.

7) Excel stores strings in Unicode UTF16. (normally 2 bytes per char). Upper to lower case 
   equivalence is much more involved with Unicode compared to ASCII.

         By choosing case insensitive; string compares in the binary search functions consume much
         larger amounts of processor time than plain binary data. Compares in these circumstances
         need to be kept to a minimum.

8) The ret variable in TRI needs to be a signed integer. The array list will also normally be
   signed positive integers but a workaround could be found for an array of unsigned integers
   provided the ret variable is signed. (such as using the 8 byte long int and typecasting in C++. 
   This would also remove the threat of overflow)

9) Some languages may require a user-defined function to be written to get the 3 state values.

-----------------------------------------------------------------------------------------------------------

#### Hardware/Software: Windows 10 64bit, Excel 2007 32bit, 250gig SSD hard disk, Intel I5 7500, 4 Gig memory.

-----------------------------------------------------------------------------------------------------------

## TESTS ON THREE TYPES OF BINARY SEARCH FUNCTIONS

 Table shows overall time in seconds to do test, searches per second (after initialise of array) and
 average number of compares needed to obtain each target value.

### 1 million array string 58-60 char CASE INSENSITIVE
### 10 million random searches

| hello MONO                    | hello STANDARD                    | hello TRI                   |
| ----------------------------- | --------------------------------- | --------------------------- |
| 54.45 secs                    | 82.79 secs                        | 50.02 secs                  |
| 183,644 searches per sec      | 120,783 searches per sec          | 199,938 searches per sec    |
| 21 average compares           | 36.9 average compares             | 18.95 average compares      |


### 1 million array string 58-60 char CASE SENSITIVE
### 10 million random searches

| hello MONO                    | hello  STANDARD                   | hello TRI                   |
| ----------------------------- | --------------------------------- | --------------------------- |
| 28.57 secs                    | 37.51 secs                        | 26.68 secs                  |
| 350,062 searches per sec      | 266,611 searches per sec          | 374,762 searches per sec    |
| 21 average compares           | 36.9 average compares             | 18.95 average compares      |


### 9 million array 32bit signed positive integer
### 9 million random searches

| hello MONO                    | hello STANDARD                    | hello TRI                   |
| ----------------------------- | --------------------------------- | --------------------------- |
| 12.17 secs                    | 11.19 secs                        | 10.84 secs                  |
| 739,409 searches per sec      | 804,188 searches per sec          | 830,570 searches per sec    |
| 25 average compares           | 43.27 average compares            | 22.14 average compares      |

---------------------------------------------------------------------------------------------------



### CONCLUSION


 1) Comparison between a target value and an array of integers in memory has **little impact** on CPU process
    time when dealing with 32 bit numbers and a 32/64 bit cpu with cache to the overall speed of searching.

 2) The STANDARD binary search is **hardly dented** for speed in integer searches despite having to do almost
    twice as many compares as TRI

 3) Therefore don't assume it's the compare against the array list that is a speed bottleneck until it's timed.
    (OR memory access time looked up in CPU reference manual)

 4) Its text searches where the compare count per search **starts to matter**. It appears to be calling a series
    of time-consuming functions to carry out the compare.

 5) A case sensitive search is about twice as fast as case insensitive.
 
---------------------------------------------------------------------------------------------------
 
### Have a nice day!
.
.
.
 


