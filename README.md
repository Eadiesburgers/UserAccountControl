# UserAccountControl

So you want to convert the retrieved UserAccountControl attribute integer into the documented text attribute value. You want to do this without building a monstrous lookup list, which would need to contain every possible permutation. 

Cooked in an oven of Notepad.exe, this VBScript converts the supplied integer response into a 32 bit (4 byte word) binary value and then displays the calculated flag. 
Additionally it will display the bit location and the associated Microsoft documented text response. The output will also visually explain how the outcome was determined and in doing so easily show how such attribute values are calculated.

Pass no params to be displayed the help text.
