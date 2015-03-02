# snippets
Random snippets

## VB/UnsignedAdd
Allows to add a (signed) increment to a LongPtr, treating the latter as unsigned. Useful for pointer arithmetic.
Should work for both 32 and 64 bit Office.

### Known issues
* overflow for UnsignedAdd(CLngPtr("-9,223,372,036,854,775,808"),CLngPtr("-9,223,372,036,854,775,808")) 