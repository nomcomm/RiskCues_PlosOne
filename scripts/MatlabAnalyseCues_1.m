% MATLAB Analyse Cues
clear all; close all; clc;
cd('D:\03_PROJEKTE\CUES\DATENSATZ')
addpath(genpath('D:\03_PROJEKTE\CUES\Matlab'));

filename = 'DATENSATZ_24-11-2014.xlsx';
[status,sheets] = xlsfinfo(filename);

% CUE-Namen einlesen
xlRange = 'A5:B588';
for i=2:5  %the 4 main indiv. datatables
    [ndata, text(i-1).text, alldata] = xlsread(filename,sheets{i}, xlRange );
end; clear xlRange ndata alldata; 
onsets = 1:8:584; iter = length(onsets);
for i = 1:4
    for j = 1: iter
        sheet(i).cue(j).names1 = text(i).text{(onsets(j)+7),1};
        sheet(i).cue(j).names2 = text(i).text{(onsets(j)+7),2};
        %sheet(i).cue(j).names3 = newcuenames{j};
        %sheet(i).cue(j).names4 = newcuenames2{j};        
    end;
end; clear i iter j onsets 

% CUE-Daten einlesen
% Range der Ratings, jeweils 8 Urteiler pro Cue, 73 Cues hintereinander d.h. 5:588 ((588-4)/8)=73, je
% 60 Bilder, d.h. C:BJ
xlRange = 'C5:BJ588';
% Cue-Ratings für die 4 verschiedenen Teile einlesen
for i=2:5
    sheet(i-1).data = xlsread(filename,sheets{i}, xlRange );
end; clear xlRange; 
clear i j onsets

% Bestimme die Punkte, bei denen die Ratings der insge. 73 Cues beginnen und
% lies alles in eine gute Struktur ein
onsets = 1:8:584; iter = length(onsets)
for i = 1:4
    for j = 1: iter
        sheet(i).cue(j).data = sheet(i).data(onsets(j):(onsets(j)+7),:);
    end;
end;
clear i j onsets

%%% CUE-Abstractheit einlesen
xlRange = 'C2:C74';
cueabst = xlsread(filename,'AbstractCue', xlRange );
for i=1:4
    for j=1:length(cueabst)
       sheet(i).cue(j).abstract = cueabst(j) 
    end;
end; clear cueabst i j 


%%% CUE-Sorting-Namen einlesen
xlRange = 'H2:I74';
[num,txt,raw]  = xlsread(filename,'CUESorting', xlRange );
for i=1:4
    for j=1:length(txt)
       sheet(i).cue(j).names3 = txt(j,1); 
       sheet(i).cue(j).names4 = txt(j,2);       
    end;
end; clear txt raw num i j 

%%%Inversionsvector einlesen
xlRange = 'D2:D74';
inversionvector  = (xlsread(filename,'CUESorting', xlRange ))';
    
% Berechne die ICCs aller Cues und die gemittelten Cues, ggf. Invertierung
% der Items für bessere Lesbarkeit/Interpretierbarkeit
sheet(3).cue(3).data = nan(8,60); %frauen bartwuchs auf 0 setzen.
sheet(4).cue(3).data = nan(8,60);
for i = 1:4
    for j = 1: iter
        M = sheet(i).cue(j).data'; 
        sheet(i).cue(j).means = nanmean(sheet(i).cue(j).data',2);  % berechne Mittelwerte
        if (inversionvector(j) == -1) %falls das Item invertiert werden muss ...
            sheet(i).cue(j).means = 8 - sheet(i).cue(j).means; 
        end;            
        X = M(~any(isnan(M),2),:); % nimm nur die Fälle, die vollst. Daten haben / könnte ggf. durch missing-value procedure ersetzt werden
        r_alpha_cues(i,j) = ICC(X, 'C-k');  
        r_icc_cues(i,j) = ICC(X, 'A-k');  
    end;
end;
%cue_reliabilities(1,:) = sortrows((nanmean(r_icc_cues(1:2,:),1)'),-1)';
%cue_reliabilities(2,:) = sortrows((nanmean(r_icc_cues(3:4,:),1)'),-1)';
clear i j X M iter

% KRITERIUMSURTEILE einlesen
% Range der Ratings, jeweils 40 Urteiler pro Bild, 120 Bilder 
xlRange = 'B2:AO121';
% Ratings für HIV (MüF, FüM), Trust (MüF, FüM), Health (MüF, FüM), Attract (MüF, FüM)
crit(1).data = xlsread(filename,'Risk-MüF', xlRange );
crit(2).data = xlsread(filename,'Risk-FüM', xlRange );
crit(3).data = xlsread(filename,'Trust-MüF', xlRange );
crit(4).data = xlsread(filename,'Trust-FüM', xlRange );
crit(5).data = xlsread(filename,'Health-MüF', xlRange );
crit(6).data = xlsread(filename,'Health-FüM', xlRange );
crit(7).data = xlsread(filename,'Att-MüF', xlRange );
crit(8).data = xlsread(filename,'Att-FüM', xlRange ); clear xlRange status filename sheets;

for i=1:8
    M = crit(i).data;
    crit(i).means = mean(M,2);
    r_alpha_crit(i) = ICC(M, 'C-k');
    r_icc_crit(i) = ICC(M, 'A-k');  
end; clear i M text;


% SORT EVERYTHING TOGETHER
% FRAUENBILDER
% #   BildNr  HIV TRUST   HEALTH  ATTRAC  MCUE1   MCUE2   ...     MCUE73
% 1   1  x   x       x       x       x       x       x       x
% 2   2  x   x       x       x       x       x       x       x
% .   .  x   x       x       x       x       x       x       x
% .   .  x   x       x       x       x       x       x       x
% .   .  x   x       x       x       x       x       x       x
% 60  60  x   x       x       x       x       x       x       x
% dasselbe für 61 - 120 sowie MAENNERBILDER
cuecategory = {'Bildnr',
                 'HIV Risk',
                 'Trust',
                 'Health',
                 'Attract'};
             for i=1:73
                 cuecategory{i+5} = sheet(1).cue(i).names1(1:(end-3));
             end;
cuenumber = {'Bildnr',
                 'HIV Risk',
                 'Trust',
                 'Health',
                 'Attract'};
             for i=1:73
                 cuenumber{i+5} = strcat('CUE',num2str(i));
             end;  
cueitem_german = {'Bildnr',
                 'HIV Risk',
                 'Trust',
                 'Health',
                 'Attract'};
             for i=1:73
                 cueitem_german{i+5} = sheet(1).cue(i).names2;
             end;   
cuefull = {'Bildnr',
                 'HIV Risk',
                 'Trust',
                 'Health',
                 'Attract'};
             for i=1:73
                 cuefull{i+5} = strcat( 'CUE', num2str(i), '_', sheet(1).cue(i).names1(1:(end-3)),'_', char(sheet(1).cue(i).names4)   , '_Abstract_', num2str(sheet(1).cue(i).abstract));
             end;           
cueitem = {'Bildnr',
                 'HIV Risk',
                 'Trust',
                 'Health',
                 'Attract'};
             for i=1:73
                 cueitem{i+5} = char(sheet(1).cue(i).names4);
             end;     
             
cueabstract = {'Bildnr',
                 'HIV Risk',
                 'Trust',
                 'Health',
                 'Attract'};
             for i=1:73
                 cueabstract{i+5} =  num2str(sheet(1).cue(i).abstract);
             end;                                         
             
for i=1:4
    struct(i).cuecategory = cuecategory;    
    struct(i).cuenumber = cuenumber;
    struct(i).cueitem_german = cueitem_german;
    struct(i).cueitem = cueitem;   
    struct(i).cuefull = cuefull;       
    struct(i).cueabstract = cueabstract;
end; clear cue* i
             
%Set1  - Maenner
MAENNER_01 = [];
count = (1:60);
indices= (1:60);
relevant_sheet = 1;
MAENNER_01 = [count' indices'];
MAENNER_01 = [MAENNER_01 crit(2).means(indices) crit(4).means(indices) crit(6).means(indices) crit(8).means(indices)];  
for i=1:73
    MAENNER_01 = [MAENNER_01 sheet(relevant_sheet).cue(i).means];
end;

%Set2  - Maenner
MAENNER_02 = [];
indices= (61:120);
relevant_sheet = 2;
MAENNER_02 = [count' indices'];
MAENNER_02 = [MAENNER_02 crit(2).means(indices) crit(4).means(indices) crit(6).means(indices) crit(8).means(indices)]; %  
for i=1:73
    MAENNER_02 = [MAENNER_02 sheet(relevant_sheet).cue(i).means];
end;

%Set1  - Frauen
FRAUEN_01 = [];
indices= (1:60);
relevant_sheet = 3;
FRAUEN_01 = [count' indices'];
FRAUEN_01 = [FRAUEN_01 crit(1).means(indices) crit(3).means(indices) crit(5).means(indices) crit(7).means(indices)]; % 
for i=1:73
    FRAUEN_01 = [FRAUEN_01 sheet(relevant_sheet).cue(i).means];
end;

%Set2  - Frauen
FRAUEN_02 = [];
indices= (61:120);
relevant_sheet = 4;
FRAUEN_02 = [count' indices'];
FRAUEN_02 = [FRAUEN_02 crit(1).means(indices) crit(3).means(indices) crit(5).means(indices) crit(7).means(indices)]; % 
for i=1:73
    FRAUEN_02 = [FRAUEN_02 sheet(relevant_sheet).cue(i).means];
end; clear i  indices relevant_sheet 

struct(1).data = MAENNER_01;
struct(2).data = MAENNER_02;
struct(3).data = FRAUEN_01;
struct(4).data = FRAUEN_02;

%Ergebnisse rausschreiben 
%header = 'No	BildNr	Risk	Trust	Health	Attractiveness	C1_2	C2_0	C3_0	C4_1	C5_1	C6_1	C7_1	C8_0	C9_1	C10_1	C11_1	C12_0	C13_1	C14_0	C15_0	C16_0	C17_0	C18_2	C19_1	C20_1	C21_1	C22_0	C23_1	C24_1	C25_1	C26_0	C27_0	C28_1	C29_1	C30_1	C31_0	C32_1	C33_0	C34_0	C35_2	C36_2	C37_2	C38_1	C39_2	C40_2	C41_2	C42_2	C43_2	C44_2	C45_2	C46_2	C47_2	C48_1	C49_0	C50_1	C51_0	C52_0	C53_0	C54_0	C55_0	C56_1	C57_0	C58_0	C59_0	C60_0	C61_0	C62_0	C63_0	C64_1	C65_0	C66_0	C67_0	C68_0	C69_0	C70_0	C71_0	C72_0	C73_1';
header = [];
tabs= sprintf('\t ');
header = 'No	BildNr	Risk	Trust	Health	Attractiveness';
for i=6:78
    aaa= {struct(1).cuefull(i)};
    header = [header, char({tabs}), char( aaa{1,1}{1,1} )]; 
    %strcat(header,{tabs}, struct(1).spaltennamen4(i));
end
    
outfile = 'D:\03_PROJEKTE\CUES\Matlab\frauen01.txt';
dlmwrite(outfile,cellstr(header),'delimiter','');
dlmwrite(outfile,FRAUEN_01,'delimiter','\t','-append');

outfile = 'D:\03_PROJEKTE\CUES\Matlab\frauen02.txt';
dlmwrite(outfile,header,'delimiter','');
dlmwrite(outfile,FRAUEN_02,'delimiter','\t','-append');

outfile = 'D:\03_PROJEKTE\CUES\Matlab\maenner01.txt';
dlmwrite(outfile,header,'delimiter','');
dlmwrite(outfile,MAENNER_01,'delimiter','\t','-append');

outfile = 'D:\03_PROJEKTE\CUES\Matlab\maenner02.txt';
dlmwrite(outfile,header,'delimiter','');
dlmwrite(outfile,MAENNER_02,'delimiter','\t','-append');
clear outfile FR* MA* all* count crit i sheet

%Berechnungen
for set = 1:4
    for i=3:6 %
        currgoalvec = struct(set).data(:,i);  %3 = HIV risk, 4 = Trust 5 Health 6 Attractiveness
        for j=1:73
            struct(set).the_corr(j,i-2)=corr(currgoalvec, struct(set).data(:,j+6));
        end
    end; 
end
%plot(the_corr','DisplayName','the_corr');
% all results are there, now do nice formatting.


dataset12 = mat2dataset(   [struct(1).the_corr(:,1) struct(3).the_corr(:,1) struct(1).the_corr(:,2) struct(3).the_corr(:,2) struct(1).the_corr(:,3) struct(3).the_corr(:,3) struct(1).the_corr(:,4) struct(3).the_corr(:,4)]);
dataset1(:,1) = cell2dataset(struct(1).cuecategory(5:end));
dataset1(:,2) = cell2dataset(struct(1).cueitem(5:end));
dataset1(:,3) = cell2dataset(struct(1).cueabstract(5:end));
dataset1(:,4:11) = dataset12; clear dataset12;

dataset1.Properties.VarNames = {'Category';'CueItem';'CueAbstract';'HIVmales';'HIVfemales';'Trustmales';'Trustfemales';'Healthmales';'Healthfemales';'Attractmales';'Attractfemales';};
dataset1.Properties.ObsNames = struct(1).cuenumber(6:end)';

%dataset1 = sortrows(dataset1,'HIVmales','descend');

export(dataset1,'File','dataset1newnames.txt')






