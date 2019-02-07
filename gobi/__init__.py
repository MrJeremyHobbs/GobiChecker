import re

class parse_line:
    def __init__(self, order_line):
        self.parse_line = parse_line
        
        #skip null lines
        if "\t" not in order_line:
            self.line_is_null = True
        else:
            self.line_is_null = False
        
            # parse
            fields = order_line.split("\t")
            
            # title
            self.title  = fields[0]
            self.title_clean = re.sub('[,:;."]', '', self.title)
            self.title_clean = re.sub('[-]', ' ', self.title_clean)
            self.title_clean = re.sub('[&]', ' ', self.title_clean)
            self.title_split = self.title_clean.split(" ")
            
            self.title_parsed_array = []
            loop_counter = 0
            for t in self.title_split:
                if loop_counter < 5:
                    self.title_parsed_array.append(t)
                    loop_counter += 1
            self.title_parsed = " ".join(self.title_parsed_array)
            
            self.title_short = self.title_split[0]
            
            # author
            self.author = fields[6]
            a = self.author.split(", ")
            self.author_lastname = a[0]
            
            # editor
            self.editor = fields[7]
            e = self.editor = self.editor.split(", ")
            
            # if no author, switch to editor
            if self.author == "":
                self.author == e[0]
                self.author_lastname = e[0]
            
            # kw
            self.kw = f"{self.author_lastname} {self.title_parsed}"
            if self.author_lastname == "":
                self.kw = f"{self.title_parsed}"
            
            # isbn
            self.isbn   = fields[10]
            
            # publisher
            self.pub    = fields[8]
            p = self.pub.split(" ")
            self.pub_short = p[0]
            
            # pub year
            self.pub_year = fields[9]
            
            # binding
            self.binding = fields[11]