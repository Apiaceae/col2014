#### The script function to treat ordered distribution information
### for col 2015 checklist data
### best run on ruby 1.9.3-p547 with nokogiri 1.6.6.2
The basic step to running the script is:

1. please input you example_moss.html file to run *adjust_distribution_order.rb* to produce file with modified distribution;
2. use pandoc to convert xxx.html to docx file, with the following command;
```
pandoc -f html 1.html -t docx -o 2.docx
pandoc -f html 1.html -t docx -o 1.docx
```
3. run *replace_xxx_family.rb* file to replace family information.

####TODO
still have problem to deal with distribution include '()' or double ';' exist
example_moss.html file include some new line and enter character which resulted unexpected results(manually remove this!!)
if some distribution like this, 分布:越南(?); 云南will result gsub function error message

