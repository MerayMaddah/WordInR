# previously the package of "ReporteRs" was used to create Word documents but has been removed since
#install the package ('officer')
install.packages('officer')
library(officer)
library(dplyr)

#create an empty word file
sample_doc <- read_docx()
#next we start adding paragraphs to the empty document
sample_doc <- sample_doc %>% body_add_par("This is the first note")
sample_doc <- sample_doc %>% body_add_par("This is the second note")
sample_doc <- sample_doc %>% body_add_par("This is the third note")
View(sample_doc)

#next we can add plotted images to our Word document

#First we create a temp file to make a plot
set.seed(0)
src <- tempfile(fileext = ".png")

#create a PNG object
png(filename = src, width = 4, height = 4, units = 'in', res = 400)

#then we create a plot
plot(sample(100, 10))

#save PNG file
dev.off()

#then we add the PNG file into the Word file
sample_doc <- sample_doc %>% body_add_img(src = src, width = 4, height = 4, style = "centered")

#Finally this is how we save our Word document, we use the Print command
print(sample_doc, target = "SampleFile.docx")
