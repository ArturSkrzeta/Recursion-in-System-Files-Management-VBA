<h2>Recursion in Files Management VBA</h2>
<h3>Intro</h3>
<p>Being tried of figuring theoretical examples out of recursion application, I decided to go ahead and apply recursion for walking through folders and processing files as per need.</p>
<p>Recursion uses stack schema where each next function is being launched on the top of the previous one. Closing functions goes the opposite way - from top to bottom.</p>
<img src="images/stack_schema.JPG">
 
<h3>Demo</h3>

<p>In VBA IDE we can see the functions Call Stack - the very first on the bottom is the main one. Afterwards, we have got functions being launched subsequently on the top of each other.</p>
  
<p>As the sequence has the stack nature, the first function that ends its work up is at the very top. Another one that ends its work up is the second one that left on the top and so on till the mian function on the bottom.</p>

<img src="images/stack.JPG">
