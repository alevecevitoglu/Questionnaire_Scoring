A = xlsread(' File_Name_Here.xlsx',’Sheet Name’); %% The scale name is good here
A(:,1) = []; %% skips the first raw which represents the IDs in my version, may skip this
A(find(isnan(A)))=0; %% Missing values made 0
X0 =sum(A,2); %% Returns the total scale score for all participant, 1 column

A = xlsread('File_Name_Here.xlsx','STAI'); %% e.g., for STAI questionnaire (sheet name)
A(:,1) = []; 
mm= 5;      %% WRITE HERE: The min+max of the scale!
list = [1, 6, 7, 10, 13, 16, 19];  %% WRITE HERE: the reverse coded items with comma!
B1(:,list)=mm-A(:,list); %% Reverse Coded version
B1(find(isnan(A)))=0; 
xlswrite(‘ReverseCode_STAI.xlsx',B1); %% Use this if you need to save the reverse code as well!
X1 =sum(A,2); 

A = xlsread('File_Name_Here.xlsx','CTQ'); %% e.g., for CTQ questionnaire (sheet name)
A(:,1) = []; 
mm= 6;      %% DON’T FORGET: This is the SUM of min and max of the scale!
list = [2, 5, 7, 13, 19, 26, 28];  %% Reverse coded items
A(:,list)=mm-A(:,list); %%This is the reverse coded version
A(find(isnan(A)))=0; 
X2 =sum(A,2); 

TotalScore = [X0,X1,X2,]; %% increase the variable number as need
xlswrite(' File_Name_Here _Coded',TotalScore); %% Returns it as cws file.
%%then select all import selection and download it.
%% Warning: This is unable to write to Excel format, attempting to write file to csv format. To write to an Excel file, convert your data to a table and use writetable. 
