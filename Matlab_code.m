clc;clear;close all;

home
filename='Registration.xlsx';
[num,txt]=xlsread(filename);

%Determine the length of the txt array
len=length(txt);

% Load a blank certificate image
blankimage=imread('Certificate_Sample.png');

for i=1:len
    for j=2:2
        text_name(i,j)=txt(i,j);
    end
end

for i=1:len
    for j=3:3
        text_topic(i,j)=txt(i,j);
    end
end

% Generate a certificate for each person in the 'txt' array

for i=2:len
    text_all=[text_name(i,2) text_topic(i,3)];
    position=[800 520;720 820];
    RGB = insertText(blankimage, position, text_all, 'FontSize', 60,'BoxOpacity', 0); 
    figure;
    imshow(RGB)
    y=i-1;
    filename=['CertificateNo_' num2str(y)];
    lastf=[filename '.png'];
    saveas(gcf,lastf);
end