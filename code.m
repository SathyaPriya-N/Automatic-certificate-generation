clear all
close all 
home 

filename = 'Registration_form.xlsx'; //  excel sheet details of participants
[num,txt] = xlsread(filename);

len=length(txt)


blankimage = imread('Blank_sample.PNG');


for i=1:len //position
    for j= 2:2 
        text_names(i,j)=txt(i,j)
    end
end

for i=1:len
    for j= 3:3
      text_topic(i,j)=txt(i,j)
    end
end


for i=2:len
        text_all=[text_names(i,2) text_topic(i,3)]
        
               
                
        position =  [267,166;267,223]; //position of text to be printed
        RGB = insertText(blankimage,position,text_all,'FontSize',15,'BoxOpacity',0);
        
        
        figure
        imshow(RGB)        
        
        y=i-1
        filename=['Certificate_Topic_' num2str(y)]
        lastf=[filename '.png']
        saveas(gcf,lastf);           
       
end    
