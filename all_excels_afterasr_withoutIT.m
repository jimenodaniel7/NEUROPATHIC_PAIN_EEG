%% Script to write data from the ASR files.

% This script exports data from ALL the files in the folder. It will have to be run in the folder where the files are.
% It creates a directory with 2 important files inside: one with the PSD data of the
% patient when he has the eyes open (OE) and one with the data from when he
% has the eyes close (CE). In each file, the first thing is a checking that
% the PSD from all 9 channels is okay. After you confirm that it is fine,
% it continues to write the data in the correct file. It allso stores files
% with the PSD of each recording and an image of a plot with each PSD
% drawn.
% Please run first the bandpass_asr.m file before this one
% Notice that you will have to change the absolute routes that are witten in some lines of this file
% before running it.
%MATLAB version: 2021b

%Input to enter the number of the subject
input_subject=input("Introduce subject number: ", 's');
   
% Sampling Frequency
Fs = 256;   

% Nyquist Frequency
Fn = Fs/2; 
resolution=0.5;
noverlap=0;
nfft=round(Fs/resolution);
window=hanning(nfft);
L = 1200;
% passband
Fpass = [0.5 40]; 
        
%We write in a file all the csv data files and in other file the events
%files so we can go extracting data for each file in the following loop
dataf_read=dir(strcat(pwd,'/*-data.xlsx'));
data_names=char(dataf_read.name);
eventsf_read=dir(strcat(pwd,'/*_events.csv'));
events_names=char(eventsf_read.name);



%Folder to store the raw data matrix
if exist('C:\\Users\\lacar\\Desktop\\sujeto_10\\raw_data')~=7
   mkdir('C:\\Users\\lacar\\Desktop\\sujeto_10\\raw_data')
end


%Folder to store data after using a bandpass filter (0.5-40 Hz)
if exist('C:\\Users\\lacar\\Desktop\\sujeto_10\\bandpass')~=7
    mkdir('C:\\Users\\lacar\\Desktop\\sujeto_10\\bandpass')
end

%Folder to store de data after using ASRpruebas
if exist('C:\\Users\\lacar\\Desktop\\sujeto_10\\asr')~=7
    mkdir('C:\\Users\\lacar\\Desktop\\sujeto_10\\asr')
end

%Folder to store the PSD files
if exist('files')~=7
    mkdir('files')
end


%loop to calculate and write the informaction of each file
for d=1:length(data_names(:,1))

    kd=strfind(data_names(d,:),".");
    name_data=data_names(d,5:kd-1);
    
    ke=strfind(events_names(d,:),".");
    name_events=events_names(d,5:ke-1);

    if exist("C:\\Users\\lacar\\Desktop\\sujeto_10\\asr\\asr_"+name_data+".xlsx")==2
        [data_asr,array_channels_names] = xlsread("C:\\Users\\lacar\\Desktop\\sujeto_10\\asr\\asr_"+name_data+".xlsx");
        bb_data=data_asr;
        bb_data_t=bb_data(:,1); %time vector (in ms)
        bb_data_sintiempo=bb_data(:,2:end); %data vector (only EEG channels)
    end

   

    %events data
    csv_events=readtable(convertCharsToStrings(events_names(d,:)), 'FileType', 'text');
    events=table2array(csv_events(1:end-1,1)); %type of the event
    events_latency=table2array(csv_events(1:end-1,2));%latency (sample number) of the event
    events_duration=table2array(csv_events(1:end-1,3));%duration (sample number) of the event

    %eeg band names to write the tables to store the data
    freq_names=["alpha" "alpha peak" "low beta" "high beta" "beta" "gamma" "theta"];
    

    %VAS value multiply by 10 so it is a number in the range 0-100
    vas=table2array(csv_events(end,2))*10;
    
    %This variable gives a classification of the measures with a number between 1 and 10 giving their VAS range (0-10 is 1, 10-20 is 2, 20-30 is 3...)
    vas_range=0;
    if vas <= 10
        vas_range=1;
    elseif vas > 10 && vas<= 20
        vas_range=2;
    elseif vas > 20 && vas<= 30
        vas_range=3;
    elseif vas > 30 && vas<= 40
        vas_range=4;
    elseif vas > 40 && vas<= 50
        vas_range=5;
    elseif vas > 50 && vas<= 60
        vas_range=6;
    elseif vas > 60 && vas<= 70
        vas_range=7;
    elseif vas > 70 && vas<= 80
        vas_range=8;
    elseif vas > 80 && vas<= 90
        vas_range=9;
    else
        vas_range=10;
    end
    
  
    %positions of events in data array. This formula gives the position of the
    %closest point in the time vector to the trigger event.

    
 
    
    %position of the ce event
    if isempty(find(table2array(csv_events(1:round(length(events_latency)/2),1))=="CE"))==0
        position_ce=events_latency(find(table2array(csv_events(1:round(length(events_latency)),1))=="CE"));
    end
  
    %position of the first oe event
    if isempty(find(table2array(csv_events(1:round(length(events_latency)/2),1))=="OE"))==0
        position_oe1=events_latency(find(table2array(csv_events(1:round(length(events_latency)/2),1))=="OE"));
    elseif isempty(find(table2array(csv_events(1:round(length(events_latency)/2),1))=="OE"))==1
        if exist('position_ce')==1 && position_ce > (60*256)
            position_oe1=2*256;
        end
    end


    %position of the first pause
    if isempty(find(table2array(csv_events(1:round(length(events_latency)/4),1))=="PAUSE"))==0
        position_p1=events_latency(find(table2array(csv_events(1:round(length(events_latency)/4),1))=="PAUSE"));
    end

    %position of the first task event
    if isempty(find(table2array(csv_events(1:round(length(events_latency)/4),1))=="IMGTASK_START"))==0
        position_task=events_latency(find(table2array(csv_events(1:round(length(events_latency)/4),1))=="PAUSE"));
    end
    
    %position of the second oe event
    if isempty(find(table2array(csv_events(round(length(events_latency)/2):end,1))=="OE"))==0
        position_oe3=events_latency(find(table2array(csv_events(round(length(events_latency)/2):end,1))=="OE"));
    end
    
    %end
    positions_end=find(bb_data_t(:,1)==(bb_data_t(end-1*256)));
    position_end=positions_end(end);

    %half of the first OE event
    if exist('position_oe1')==1 && exist('position_ce')==1
        position_oe2=(position_oe1 + position_ce) /2;
    end

    %half of the second OE event
    if exist('position_oe3')==1 && exist('position_end')==1
        position_oe4=(position_oe3 + position_end) /2;
    end

    %Half oof the CE event
    if exist('position_ce')==1 && exist('position_p1')==1
        position_ce2=(position_ce + position_p1) /2;
    end

    
    if exist("C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures")~=7
        mkdir("C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures")
        mkdir("C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr")
        mkdir("C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_sinasr")
    end


    %%PSD FIGURES (ASR)%%
    %first, there is a comprobation to see if the ASR haven't cut the part
    %of the signal of each event.
    %% 1OE 
    if exist('position_oe1')==1 && exist('position_oe2')==1    
        [psd,f]=pwelch(bb_data_sintiempo(position_oe1:position_oe2,:),window,[],nfft,Fs);
        figure(1)
        for c=1:(length(array_channels_names)-1)
            subplot(1,3,1)
            if array_channels_names(c+1)=="FC3" || array_channels_names(c+1)=="C3" || array_channels_names(c+1)=="CP3" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD left channels")
                legend show
                axis([0 29 0 15])
                hold on
            end
            subplot(1,3,2)
            if array_channels_names(c+1)=="FCz" || array_channels_names(c+1)=="Cz" || array_channels_names(c+1)=="CPZ" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD central channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            if array_channels_names(c+1)=="FC4" || array_channels_names(c+1)=="C4" || array_channels_names(c+1)=="CP4" 
                
                subplot(1,3,3)
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD right channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
       
        end
        sgtitle(name_data+" 10E ASR")
        saveas(gcf,"C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_1OE_asr.jpg")
        close
    end    
 
    %% 2OE 

    if exist('position_oe2')==1 && exist('position_ce')==1    
        [psd,f]=pwelch(bb_data_sintiempo(position_oe2:position_ce,:),window,[],nfft,Fs);
         figure(2)
        for c=1:(length(array_channels_names)-1)
            subplot(1,3,1)
            if array_channels_names(c+1)=="FC3" || array_channels_names(c+1)=="C3" || array_channels_names(c+1)=="CP3" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD left channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            subplot(1,3,2)
            if array_channels_names(c+1)=="FCz" || array_channels_names(c+1)=="Cz" || array_channels_names(c+1)=="CPZ" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD central channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            if array_channels_names(c+1)=="FC4" || array_channels_names(c+1)=="C4" || array_channels_names(c+1)=="CP4" 
                
                subplot(1,3,3)
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD right channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
        end
        sgtitle(name_data+" 20E ASR")
        saveas(gcf,"C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_2OE_asr.jpg")
        close
    end    



    %% 1CE 

    if exist('position_ce')==1 && exist('position_ce2')==1
        [psd,f]=pwelch(bb_data_sintiempo(position_ce:position_ce2,:),window,[],nfft,Fs);
         figure(3)
        for c=1:(length(array_channels_names)-1)
            subplot(1,3,1)
            if array_channels_names(c+1)=="FC3" || array_channels_names(c+1)=="C3" || array_channels_names(c+1)=="CP3" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD left channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            subplot(1,3,2)
            if array_channels_names(c+1)=="FCz" || array_channels_names(c+1)=="Cz" || array_channels_names(c+1)=="CPZ" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD central channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            if array_channels_names(c+1)=="FC4" || array_channels_names(c+1)=="C4" || array_channels_names(c+1)=="CP4" 
                
                subplot(1,3,3)
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD right channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
        end
        sgtitle(name_data+" CE1 ASR")
        saveas(gcf,"C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_1CE_asr.jpg")
        close
    end    

    %% 2CE 
   
    if exist('position_ce2')==1 && exist('position_p1')==1
        [psd,f]=pwelch(bb_data_sintiempo(position_ce2:position_p1,:),window,[],nfft,Fs);
        figure(4)
        for c=1:(length(array_channels_names)-1)
            subplot(1,3,1)
            if array_channels_names(c+1)=="FC3" || array_channels_names(c+1)=="C3" || array_channels_names(c+1)=="CP3" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD left channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            subplot(1,3,2)
            if array_channels_names(c+1)=="FCz" || array_channels_names(c+1)=="Cz" || array_channels_names(c+1)=="CPZ" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD central channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            if array_channels_names(c+1)=="FC4" || array_channels_names(c+1)=="C4" || array_channels_names(c+1)=="CP4" 
                
                subplot(1,3,3)
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD right channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
        end
        sgtitle(name_data+" CE2 ASR")
        saveas(gcf,"C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_2CE_asr.jpg")
        close
    end    


    %% 3OE 
  

    if exist('position_oe3')==1 && exist('position_oe4')==1
        [psd,f]=pwelch(bb_data_sintiempo(position_oe3:position_oe4,:),window,[],nfft,Fs);
        figure(5)
        for c=1:(length(array_channels_names)-1)
            subplot(1,3,1)
            if array_channels_names(c+1)=="FC3" || array_channels_names(c+1)=="C3" || array_channels_names(c+1)=="CP3" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD left channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            subplot(1,3,2)
            if array_channels_names(c+1)=="FCz" || array_channels_names(c+1)=="Cz" || array_channels_names(c+1)=="CPZ" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD central channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            if array_channels_names(c+1)=="FC4" || array_channels_names(c+1)=="C4" || array_channels_names(c+1)=="CP4" 
                
                subplot(1,3,3)
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD right channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
        end
        sgtitle(name_data+" OE3 ASR")
        saveas(gcf,"C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_3OE_asr.jpg")
        close
    end  

    %% 4OE 

    if exist('position_oe4')==1 && exist('position_end')==1
        [psd,f]=pwelch(bb_data_sintiempo(position_oe4:position_end,:),window,[],nfft,Fs);
        figure(6)
        for c=1:(length(array_channels_names)-1)
            subplot(1,3,1)
            if array_channels_names(c+1)=="FC3" || array_channels_names(c+1)=="C3" || array_channels_names(c+1)=="CP3" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD left channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            subplot(1,3,2)
            if array_channels_names(c+1)=="FCz" || array_channels_names(c+1)=="Cz" || array_channels_names(c+1)=="CPZ" 
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD central channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
            if array_channels_names(c+1)=="FC4" || array_channels_names(c+1)=="C4" || array_channels_names(c+1)=="CP4" 
                
                subplot(1,3,3)
                plot (f (5:80), psd(5:80,c),'DisplayName',cell2mat(array_channels_names(c+1)));
                title("PSD right channels")
                axis([0 29 0 15])
                legend show
                hold on
            end
        end
        sgtitle(name_data+" OE4 ASR")
        saveas(gcf,"C:\\Users\\lacar\\Desktop\\sujeto_10\\"+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_4OE_asr.jpg")
        close
    end  

    close all

    %% Day of record

    record_day=d;

    %% Excel OE
    %the excel will have 9 worksheets (one for each channel). Inside each
    %worksheets, there will be 7 tables, one for each eeg band (alpha, 
    %beta divided in low, high and complete, gamma, thetha) plus one that show the 
    %alpha peaks in each recording. The rows of the
    %tables will be the different recordings.
    
    
    OE=["OE min 1" "OE min 2" "OE min 3" "OE4 min 4" "VAS" "VAS_RANGE" " "]; %to create the titles of the columns inn the excel
    OE_FINAL=[OE OE(1,1:4) " " OE OE OE OE OE];

    for i= 1:length(bb_data_sintiempo(1,:))
        xlswrite("files/OE.xls",mat2cell(name_data,1),cell2mat(array_channels_names(i+1)), "A"+int2str(record_day+5))
        xlswrite("files/OE.xls",OE_FINAL, cell2mat(array_channels_names(i+1)), 'B5')
        xlswrite("files/OE.xls",mat2cell(freq_names(1),1),cell2mat(array_channels_names(i+1)), 'B4') %alpha
        xlswrite("files/OE.xls",mat2cell(freq_names(2),1),cell2mat(array_channels_names(i+1)), 'I4') %alpha peak
        xlswrite("files/OE.xls",mat2cell(freq_names(3),1),cell2mat(array_channels_names(i+1)), 'N4') %high beta
        xlswrite("files/OE.xls",mat2cell(freq_names(4),1),cell2mat(array_channels_names(i+1)), 'U4') %low beta
        xlswrite("files/OE.xls",mat2cell(freq_names(5),1),cell2mat(array_channels_names(i+1)), 'AB4') %beta
        xlswrite("files/OE.xls",mat2cell(freq_names(6),1),cell2mat(array_channels_names(i+1)), 'AI4') %gamma
        xlswrite("files/OE.xls",mat2cell(freq_names(7),1),cell2mat(array_channels_names(i+1)), 'AP4') %theta
        
    end
    
    %create some empty arrays to store the MATLAB variables we create in
    %every instance of the for loop
    psd_oe1_total=[];
    potencia_oe1=[];
    psd_oe1_relative_alfa=[];
    psd_oe2_total=[];
    potencia_oe2=[];
    psd_oe2_relative_alfa=[];
    psd_oe3_total=[];
    potencia_oe3=[];
    psd_oe3_relative_alfa=[];
    psd_oe4_total=[];
    potencia_oe4=[];
    psd_oe4_relative_alfa=[];
    
    %this loop writes the data inside each table
    
    for i=1:length(bb_data_sintiempo(1,:))
       
        %psd of each minute of the OE event
        if exist('position_oe1')==1 && exist('position_oe2')==1
            bb_data_oe1=bb_data_sintiempo(position_oe1:position_oe2,:);
            [psd_oe1,f_oe1]=pwelch(bb_data_oe1(:,i),window,[],nfft,Fs); %psd
            psd_oe1_total=[psd_oe1_total psd_oe1];
        end
        if exist('position_oe2')==1 && exist('position_ce')==1
            bb_data_oe2=bb_data_sintiempo(position_oe2:position_ce,:);
            [psd_oe2,f_oe2]=pwelch(bb_data_oe2(:,i),window,[],nfft,Fs);%psd
            psd_oe2_total=[psd_oe2_total psd_oe2];
        end
        if exist('position_oe3')==1 && exist('position_oe4')==1
            bb_data_oe3=bb_data_sintiempo(position_oe3:position_oe4,:);
            [psd_oe3,f_oe3]=pwelch(bb_data_oe3(:,i),window,[],nfft,Fs);%psd
            psd_oe3_total=[psd_oe3_total psd_oe3];
        end
        if exist('position_oe4')==1 && exist('position_end')==1
            bb_data_oe4=bb_data_sintiempo(position_oe4:position_end,:);
            [psd_oe4,f_oe4]=pwelch(bb_data_oe4(:,i),window,[],nfft,Fs);%psd
            psd_oe4_total=[psd_oe4_total psd_oe4];
        end
        
        
        
            
        %MIN 1 OE

        if exist('psd_oe1')==1
            %Alpha region 
        
            % In the alpha region, we must first find the peak of the alpha
            % frequencies. Then, we define the region from the peak -2Hz to the
            % peak +2Hz. This is because not everyone has the same alpha region.
        
            pico_alpha=f(find(psd_oe1(:,1)==max(psd_oe1(13:30,1))));
        
            %psd_sum_alpha_oe1 = sum(psd_oe1(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %sum of psd values for delta frequency region
            psd_int_alpha_oe1 = trapz(psd_oe1(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_alpha_oe1,cell2mat(array_channels_names(i+1)),"B"+int2str(record_day+5));
            xlswrite("files/OE.xls",pico_alpha,cell2mat(array_channels_names(i+1)),"I"+int2str(record_day+5));
    
            poe1=trapz(psd_oe1(5:58,1));
            psd_alpha_r_oe1=psd_int_alpha_oe1/poe1;
            psd_oe1_relative_alfa=[psd_oe1_relative_alfa psd_alpha_r_oe1];
            potencia_oe1=[potencia_oe1 poe1];

            %Low-beta region (13-20 Hz)
            %psd_sum_low_beta_oe1 = sum(psd_oe1(find(f==13):find(f==20),1)); %sum of psd values for delta frequency region
            psd_int_low_beta_oe1 = trapz(psd_oe1(find(f==13):find(f==20),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_low_beta_oe1,cell2mat(array_channels_names(i+1)),"N"+int2str(record_day+5));
            
            %High beta region (20-30 Hz)
            %psd_sum_high_beta_oe1 = sum(psd_oe1(find(f==20):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_high_beta_oe1 = trapz(psd_oe1(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
            
            xlswrite("files/OE.xls",psd_int_high_beta_oe1,cell2mat(array_channels_names(i+1)),"U"+int2str(record_day+5));
        
            %beta region (13-30 Hz)
            %psd_sum_beta_oe1 = sum(psd_oe1(find(f==13):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_beta_oe1 = trapz(psd_oe1(find(f==13):find(f==30),1)); %integral of psd values for delta frequency region
            
            xlswrite("files/OE.xls",psd_int_beta_oe1,cell2mat(array_channels_names(i+1)),"AB"+int2str(record_day+5));
        
            %Gamma region (30-80 Hz but we use only until 40)
            %psd_sum_gamma_oe1 = sum(psd_oe1(find(f==30):find(f==40),1)); %sum of psd values for delta frequency region
            psd_int_gamma_oe1 = trapz(psd_oe1(find(f==30):find(f==40),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_gamma_oe1,cell2mat(array_channels_names(i+1)),"AI"+int2str(record_day+5));
        
            
            %Theta region (4-8 Hz)
            %psd_sum_theta_oe1 = sum(psd_oe1(find(f==4):find(f==8),1)); %sum of psd values for delta frequency region
            psd_int_theta_oe1 = trapz(psd_oe1(find(f==4):find(f==8),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_theta_oe1,cell2mat(array_channels_names(i+1)),"AP"+int2str(record_day+5));
        
        end

        %MIN 2 OE

        if exist('psd_oe2')==1

            %Alpha region (8-12 Hz)
        
        
            pico_alpha=f(find(psd_oe2(:,1)==max(psd_oe2(13:30,1))));
        
            %psd_sum_alpha_oe2 = sum(psd_oe2(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %sum of psd values for delta frequency region
            psd_int_alpha_oe2 = trapz(psd_oe2(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_alpha_oe2,cell2mat(array_channels_names(i+1)),"C"+int2str(record_day+5));
            xlswrite("files/OE.xls",pico_alpha,cell2mat(array_channels_names(i+1)),"J"+int2str(record_day+5));
            
            p_oe2=trapz(psd_oe2(5:58,1));
            psd_alpha_r_oe2=psd_int_alpha_oe2/p_oe2;
            psd_oe2_relative_alfa=[psd_oe2_relative_alfa psd_alpha_r_oe2];
            potencia_oe2=[potencia_oe2 p_oe2];

            %Low beta region (13-20 Hz)
            %psd_sum_low_beta_oe2 = sum(psd_oe2(find(f==13):find(f==20),1)); %sum of psd values for delta frequency region
            psd_int_low_beta_oe2 = trapz(psd_oe2(find(f==13):find(f==20),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_low_beta_oe2,cell2mat(array_channels_names(i+1)),"O"+int2str(record_day+5));
        
            %High beta region (20-30 Hz)
            %psd_sum_high_beta_oe2 = sum(psd_oe2(find(f==20):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_high_beta_oe2 = trapz(psd_oe2(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_high_beta_oe2,cell2mat(array_channels_names(i+1)),"V"+int2str(record_day+5));
        
            %beta region (13-30 Hz)
            %psd_sum_beta_oe2 = sum(psd_oe2(find(f==13):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_beta_oe2 = trapz(psd_oe2(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_beta_oe2,cell2mat(array_channels_names(i+1)),"AC"+int2str(record_day+5));
        
            %Gamma region (30-80 Hz but we use only until 40)
            %psd_sum_gamma_oe2 = sum(psd_oe2(find(f==30):find(f==40),1)); %sum of psd values for delta frequency region
            psd_int_gamma_oe2 = trapz(psd_oe2(find(f==30):find(f==40),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_gamma_oe2,cell2mat(array_channels_names(i+1)),"AJ"+int2str(record_day+5));
        
            
            %Theta region (4-8 Hz)
            %psd_sum_theta_oe2 = sum(psd_oe2(find(f==4):find(f==8),1)); %sum of psd values for delta frequency region
            psd_int_theta_oe2 = trapz(psd_oe2(find(f==4):find(f==8),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_theta_oe2,cell2mat(array_channels_names(i+1)),"AQ"+int2str(record_day+5));
        end
    
        %MIN 3 OE

        if exist('psd_oe3')==1

            %Alpha region (8-12 Hz)
        
            pico_alpha=f(find(psd_oe3(:,1)==max(psd_oe3(13:30,1))));
        
            %psd_sum_alpha_oe3 = sum(psd_oe3(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %sum of psd values for delta frequency region
            psd_int_alpha_oe3 = trapz(psd_oe3(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_alpha_oe3,cell2mat(array_channels_names(i+1)),"D"+int2str(record_day+5));
            xlswrite("files/OE.xls",pico_alpha,cell2mat(array_channels_names(i+1)),"K"+int2str(record_day+5));
      
            p_oe3=trapz(psd_oe3(5:58,1));
            psd_alpha_r_oe3=psd_int_alpha_oe3/p_oe3;
            psd_oe3_relative_alfa=[psd_oe3_relative_alfa psd_alpha_r_oe3];
            potencia_oe3=[potencia_oe3 p_oe3];

            %Low beta region (13-20 Hz)
            %psd_sum_low_beta_oe3 = sum(psd_oe3(find(f==13):find(f==20),1)); %sum of psd values for delta frequency region
            psd_int_low_beta_oe3 = trapz(psd_oe3(find(f==13):find(f==20),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_low_beta_oe3,cell2mat(array_channels_names(i+1)),"P"+int2str(record_day+5));
          
            %High beta region (20-30 Hz)
            %psd_sum_high_beta_oe3 = sum(psd_oe3(find(f==20):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_high_beta_oe3 = trapz(psd_oe3(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_high_beta_oe3,cell2mat(array_channels_names(i+1)),"W"+int2str(record_day+5));
        
            %beta region (13-30 Hz)
            %psd_sum_beta_oe3 = sum(psd_oe3(find(f==13):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_beta_oe3 = trapz(psd_oe3(find(f==13):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_beta_oe3,cell2mat(array_channels_names(i+1)),"AD"+int2str(record_day+5));
        
            %Gamma region (30-80 Hz but we use only until 40)
            %psd_sum_gamma_oe3 = sum(psd_oe3(find(f==30):find(f==40),1)); %sum of psd values for delta frequency region
            psd_int_gamma_oe3 = trapz(psd_oe3(find(f==30):find(f==40),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_gamma_oe3,cell2mat(array_channels_names(i+1)),"AK"+int2str(record_day+5));
        
        
            %Theta region (4-8 Hz)
            %psd_sum_theta_oe3 = sum(psd_oe3(find(f==4):find(f==8),1)); %sum of psd values for delta frequency region
            psd_int_theta_oe3 = trapz(psd_oe3(find(f==4):find(f==8),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_theta_oe3,cell2mat(array_channels_names(i+1)),"AR"+int2str(record_day+5));
        end

        %MIN 4 OE

        if exist('psd_oe4')==1
    
            %Alpha region (8-12 Hz)
        
            pico_alpha=f(find(psd_oe4(:,1)==max(psd_oe4(13:30,1))));
        
            %psd_sum_alpha_oe4 = sum(psd_oe4(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %sum of psd values for delta frequency region
            psd_int_alpha_oe4 = trapz(psd_oe4(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_alpha_oe4,cell2mat(array_channels_names(i+1)),"E"+int2str(record_day+5));
            xlswrite("files/OE.xls",pico_alpha,cell2mat(array_channels_names(i+1)),"L"+int2str(record_day+5));
    
            p_oe4=trapz(psd_oe4(5:58,1));
            psd_alpha_r_oe4=psd_int_alpha_oe4/p_oe4;
            psd_oe4_relative_alfa=[psd_oe4_relative_alfa psd_alpha_r_oe4];
            potencia_oe4=[potencia_oe4 p_oe4];  %total power of psd

            %Low beta region (13-20 Hz)
            %psd_sum_low_beta_oe4 = sum(psd_oe4(find(f==13):find(f==20),1)); %sum of psd values for delta frequency region
            psd_int_low_beta_oe4 = trapz(psd_oe4(find(f==13):find(f==20),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_low_beta_oe4,cell2mat(array_channels_names(i+1)),"Q"+int2str(record_day+5));
            
            %High beta region (20-30 Hz)
            %psd_sum_high_beta_oe4 = sum(psd_oe4(find(f==20):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_high_beta_oe4 = trapz(psd_oe4(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_high_beta_oe4,cell2mat(array_channels_names(i+1)),"X"+int2str(record_day+5));
        
            %beta region (13-30 Hz)
            %psd_sum_beta_oe4 = sum(psd_oe4(find(f==13):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_beta_oe4 = trapz(psd_oe4(find(f==13):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_beta_oe4,cell2mat(array_channels_names(i+1)),"AE"+int2str(record_day+5));
        
            %Gamma region (30-80 Hz but we use only until 40)
            %psd_sum_gamma_oe4 = sum(psd_oe4(find(f==30):find(f==40),1)); %sum of psd values for delta frequency region
            psd_int_gamma_oe4 = trapz(psd_oe4(find(f==30):find(f==40),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/OE.xls",psd_int_gamma_oe4,cell2mat(array_channels_names(i+1)),"AL"+int2str(record_day+5));
        
        
            %Theta region (4-8 Hz)
            %psd_sum_theta_oe4 = sum(psd_oe4(find(f==4):find(f==8),1)); %sum of psd values for delta frequency region
            psd_int_theta_oe4 = trapz(psd_oe4(find(f==4):find(f==8),1)); %integral of psd values for delta frequency region
       
            xlswrite("files/OE.xls",psd_int_theta_oe4,cell2mat(array_channels_names(i+1)),"AS"+int2str(record_day+5));
        end
    
        
        %write VAS and VAS range in each table
        xlswrite("files/OE.xls",vas,cell2mat(array_channels_names(i+1)),"F"+int2str(record_day+5));
        xlswrite("files/OE.xls",vas_range,cell2mat(array_channels_names(i+1)),"G"+int2str(record_day+5));
    
        xlswrite("files/OE.xls",vas,cell2mat(array_channels_names(i+1)),"R"+int2str(record_day+5));
        xlswrite("files/OE.xls",vas_range,cell2mat(array_channels_names(i+1)),"S"+int2str(record_day+5));
    
        xlswrite("files/OE.xls",vas,cell2mat(array_channels_names(i+1)),"Y"+int2str(record_day+5));
        xlswrite("files/OE.xls",vas_range,cell2mat(array_channels_names(i+1)),"Z"+int2str(record_day+5));
    
        xlswrite("files/OE.xls",vas,cell2mat(array_channels_names(i+1)),"AF"+int2str(record_day+5));
        xlswrite("files/OE.xls",vas_range,cell2mat(array_channels_names(i+1)),"AG"+int2str(record_day+5));
    
        xlswrite("files/OE.xls",vas,cell2mat(array_channels_names(i+1)),"AM"+int2str(record_day+5));
        xlswrite("files/OE.xls",vas_range,cell2mat(array_channels_names(i+1)),"AN"+int2str(record_day+5));
    
        xlswrite("files/OE.xls",vas,cell2mat(array_channels_names(i+1)),"AT"+int2str(record_day+5));
        xlswrite("files/OE.xls",vas_range,cell2mat(array_channels_names(i+1)),"AU"+int2str(record_day+5));
    
        %write the subject number
        if exist('input_subject')==1
            xlswrite("files/OE.xls",mat2cell(input_subject,1),cell2mat(array_channels_names(i+1)),"A1");
        end
    end
    if exist("f")
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("Frequency(Hz)",1),"PSD","A2");
        xlswrite("files/"+name_data+"_PSD_OE.xls",f,"PSD","A3");
    end
    if ne(isempty(psd_oe1_total),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("OE MIN 1",1),"PSD","B1");
        xlswrite("files/"+name_data+"_PSD_OE.xls",array_channels_names(2:end),"PSD","B2");
        xlswrite("files/"+name_data+"_PSD_OE.xls",psd_oe1_total,"PSD","B3");
    end
    if ne(isempty(psd_oe2_total),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("OE MIN 2",1),"PSD","L1");
        xlswrite("files/"+name_data+"_PSD_OE.xls",array_channels_names(2:end),"PSD","L2");
        xlswrite("files/"+name_data+"_PSD_OE.xls",psd_oe2_total,"PSD","L3");
    end
    if ne(isempty(psd_oe3_total),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("OE MIN 3",1),"PSD","V1");
        xlswrite("files/"+name_data+"_PSD_OE.xls",array_channels_names(2:end),"PSD","V2");
        xlswrite("files/"+name_data+"_PSD_OE.xls",psd_oe3_total,"PSD","V3");
    end
    if ne(isempty(psd_oe4_total),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("OE MIN 4",1),"PSD","AF1");
        xlswrite("files/"+name_data+"_PSD_OE.xls",array_channels_names(2:end),"PSD","AF2");
        xlswrite("files/"+name_data+"_PSD_OE.xls",psd_oe4_total,"PSD","AF3");
    end

    if ne(isempty(potencia_oe1),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("Power",1),"PSD","B265");
        xlswrite("files/"+name_data+"_PSD_OE.xls",potencia_oe1,"PSD","B266");
    end
    if ne(isempty(potencia_oe2),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("Power",1),"PSD","L265");
        xlswrite("files/"+name_data+"_PSD_OE.xls",potencia_oe2,"PSD","L266");
    end
    if ne(isempty(potencia_oe3),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("Power",1),"PSD","V265");
        xlswrite("files/"+name_data+"_PSD_OE.xls",potencia_oe3,"PSD","V266");
    end
    if ne(isempty(potencia_oe4),1)
        xlswrite("files/"+name_data+"_PSD_OE.xls",mat2cell("Power",1),"PSD","AF265");
        xlswrite("files/"+name_data+"_PSD_OE.xls",potencia_oe4,"PSD","AF266");
    end
    
        

    %% Excel CE
    %the excel will have 9 worksheets (one for each channel). Inside each
    %worksheets, there will be 7 tables, one for each eeg band (alpha, 
    %beta divided in low and high, gamma, thetha) plus one that show the 
    %alpha peaks in each recording. The rows of the
    %tables will be the different recordings.
    
    CE=["CE min 1" "CE min 2" "VAS" "VAS_RANGE" " "]; %to create the titles of the columns inn the excel
    CE_FINAL=[CE CE(1,1:2) " " CE CE CE CE CE];
 
    for i=1:length(bb_data_sintiempo(1,:))
        xlswrite("files/CE.xls",mat2cell(name_data,1),cell2mat(array_channels_names(i+1)), "A"+int2str(record_day+5))
        xlswrite("files/CE.xls",CE_FINAL,cell2mat(array_channels_names(i+1)), 'B5')
        xlswrite("files/CE.xls",mat2cell(freq_names(1),1),cell2mat(array_channels_names(i+1)), 'B4')
        xlswrite("files/CE.xls",mat2cell(freq_names(2),1),cell2mat(array_channels_names(i+1)), 'G4')
        xlswrite("files/CE.xls",mat2cell(freq_names(3),1),cell2mat(array_channels_names(i+1)), 'J4')
        xlswrite("files/CE.xls",mat2cell(freq_names(4),1),cell2mat(array_channels_names(i+1)), 'O4')
        xlswrite("files/CE.xls",mat2cell(freq_names(5),1),cell2mat(array_channels_names(i+1)), 'T4')
        xlswrite("files/CE.xls",mat2cell(freq_names(6),1),cell2mat(array_channels_names(i+1)), 'Y4')
        xlswrite("files/CE.xls",mat2cell(freq_names(7),1),cell2mat(array_channels_names(i+1)), 'AD4')
  
    end
    
    
    psd_ce1_total=[];
    potencia_ce1=[];
    psd_ce2_total=[];
    potencia_ce2=[];
    psd_ce1_relative_alfa=[];
    psd_ce2_relative_alfa=[];


    %this loop writes the data inside each table
    
    for i=1:length(bb_data_sintiempo(1,:))
    

        %psd of each minute of the OE event
        if exist('position_ce')==1 && exist('position_ce2')==1
            bb_data_ce1=bb_data_sintiempo(position_ce:position_ce2,:);
            [psd_ce1,f_ce1]=pwelch(bb_data_ce1(:,i),window,[],nfft,Fs); %psd
            psd_ce1_total=[psd_ce1_total psd_ce1];
        end
        if exist('position_ce2')==1 && exist('position_p1')==1
            bb_data_ce2=bb_data_sintiempo(position_ce2:position_p1,:);
            [psd_ce2,f_ce2]=pwelch(bb_data_ce2(:,i),window,[],nfft,Fs);%psd
            psd_ce2_total=[psd_ce2_total psd_ce2];
        end


        
    
        %MIN 1 CE

        if exist('psd_ce1')==1

            %Alpha region 
            pico_alpha=f(find(psd_ce1(:,1)==max(psd_ce1(13:30,1))));
            %psd_sum_alpha_ce1 = sum(psd_ce1(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %sum of psd values for delta frequency region
            psd_int_alpha_ce1 = trapz(psd_ce1(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_alpha_ce1,cell2mat(array_channels_names(i+1)),"B"+int2str(record_day+5));
            xlswrite("files/CE.xls",pico_alpha,cell2mat(array_channels_names(i+1)),"G"+int2str(record_day+5));
           
            pce1=trapz(psd_ce1(5:58,1));
            psd_alpha_r_ce1=psd_int_alpha_ce1/pce1;
            psd_ce1_relative_alfa=[psd_ce1_relative_alfa psd_alpha_r_ce1];% relative alpha
            potencia_ce1=[potencia_ce1 pce1]; %total power of psd



            %Low beta region (13-20 Hz)
            %psd_sum_low_beta_ce1 = sum(psd_ce1(find(f==13):find(f==20),1)); %sum of psd values for delta frequency region
            psd_int_low_beta_ce1 = trapz(psd_ce1(find(f==13):find(f==20),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_low_beta_ce1,cell2mat(array_channels_names(i+1)),"J"+int2str(record_day+5));
         
            %High beta region (20-30 Hz)
            %psd_sum_high_beta_ce1 = sum(psd_ce1(find(f==20):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_high_beta_ce1 = trapz(psd_ce1(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_high_beta_ce1,cell2mat(array_channels_names(i+1)),"O"+int2str(record_day+5));
        
         
            %beta region (13-30 Hz)
            %psd_sum_beta_ce1 = sum(psd_ce1(find(f==13):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_beta_ce1 = trapz(psd_ce1(find(f==13):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_beta_ce1,cell2mat(array_channels_names(i+1)),"T"+int2str(record_day+5));
        
            %Gamma region (30-80 Hz but we use only until 40)
            %psd_sum_gamma_ce1 = sum(psd_ce1(find(f==30):find(f==40),1)); %sum of psd values for delta frequency region
            psd_int_gamma_ce1 = trapz(psd_ce1(find(f==30):find(f==40),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_gamma_ce1,cell2mat(array_channels_names(i+1)),"Y"+int2str(record_day+5));
        
           
            %Theta region (4-8 Hz)
            %psd_sum_theta_ce1 = sum(psd_ce1(find(f==4):find(f==8),1)); %sum of psd values for delta frequency region
            psd_int_theta_ce1 = trapz(psd_ce1(find(f==4):find(f==8),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_theta_ce1,cell2mat(array_channels_names(i+1)),"AD"+int2str(record_day+5));
        end
    
        %MIN 2 CE

        if exist('psd_ce2')==1

            %Alpha region (8-12 Hz)
        
            pico_alpha=f(find(psd_ce2(:,1)==max(psd_ce2(13:30,1))));
            %psd_sum_alpha_ce2 = sum(psd_ce2(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %sum of psd values for delta frequency region
            psd_int_alpha_ce2 = trapz(psd_ce2(find(f==pico_alpha-2):find(f==pico_alpha+2),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_alpha_ce2,cell2mat(array_channels_names(i+1)),"C"+int2str(record_day+5));
            xlswrite("files/CE.xls",pico_alpha,cell2mat(array_channels_names(i+1)),"H"+int2str(record_day+5));
            
            pce2=trapz(psd_ce2(5:58,1));
            psd_alpha_r_ce2=psd_int_alpha_ce2/pce2;
            psd_ce2_relative_alfa=[psd_ce2_relative_alfa psd_alpha_r_ce2]; %relative alpha
            potencia_ce2=[potencia_ce2 pce2]; %total power of psd


            %Low Beta region (13-20 Hz)
            %psd_sum_low_beta_ce2 = sum(psd_ce2(find(f==13):find(f==20),1)); %sum of psd values for delta frequency region
            psd_int_low_beta_ce2 = trapz(psd_ce2(find(f==13):find(f==20),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_low_beta_ce2,cell2mat(array_channels_names(i+1)),"K"+int2str(record_day+5));
            
            %High beta region (20-30 Hz)
            %psd_sum_high_beta_ce2 = sum(psd_ce2(find(f==20):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_high_beta_ce2 = trapz(psd_ce2(find(f==20):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_high_beta_ce2,cell2mat(array_channels_names(i+1)),"P"+int2str(record_day+5));
            
            %beta region (13-30 Hz)
            %psd_sum_beta_ce2 = sum(psd_ce2(find(f==13):find(f==30),1)); %sum of psd values for delta frequency region
            psd_int_beta_ce2 = trapz(psd_ce2(find(f==13):find(f==30),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_beta_ce2,cell2mat(array_channels_names(i+1)),"U"+int2str(record_day+5));
            
            %Gamma region (30-80 Hz but we use only until 40)
            %psd_sum_gamma_ce2 = sum(psd_ce2(find(f==30):find(f==40),1)); %sum of psd values for delta frequency region
            psd_int_gamma_ce2 = trapz(psd_ce2(find(f==30):find(f==40),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_gamma_ce2,cell2mat(array_channels_names(i+1)),"Z"+int2str(record_day+5));
        
            %Theta region (4-8 Hz)
            %psd_sum_theta_ce2 = sum(psd_ce2(find(f==4):find(f==8),1)); %sum of psd values for delta frequency region
            psd_int_theta_ce2 = trapz(psd_ce2(find(f==4):find(f==8),1)); %integral of psd values for delta frequency region
        
            xlswrite("files/CE.xls",psd_int_theta_ce2,cell2mat(array_channels_names(i+1)),"AE"+int2str(record_day+5));
        
        end

       
    
        %write in each table of the excel the VAS and VAS range
        xlswrite("files/CE.xls",vas,cell2mat(array_channels_names(i+1)),"D"+int2str(record_day+5));
        xlswrite("files/CE.xls",vas_range,cell2mat(array_channels_names(i+1)),"E"+int2str(record_day+5));
    
        xlswrite("files/CE.xls",vas,cell2mat(array_channels_names(i+1)),"L"+int2str(record_day+5));
        xlswrite("files/CE.xls",vas_range,cell2mat(array_channels_names(i+1)),"M"+int2str(record_day+5));
    
        xlswrite("files/CE.xls",vas,cell2mat(array_channels_names(i+1)),"Q"+int2str(record_day+5));
        xlswrite("files/CE.xls",vas_range,cell2mat(array_channels_names(i+1)),"R"+int2str(record_day+5));
    
        xlswrite("files/CE.xls",vas,cell2mat(array_channels_names(i+1)),"V"+int2str(record_day+5));
        xlswrite("files/CE.xls",vas_range,cell2mat(array_channels_names(i+1)),"W"+int2str(record_day+5));
    
        xlswrite("files/CE.xls",vas,cell2mat(array_channels_names(i+1)),"AA"+int2str(record_day+5));
        xlswrite("files/CE.xls",vas_range,cell2mat(array_channels_names(i+1)),"AB"+int2str(record_day+5));
    
        xlswrite("files/CE.xls",vas,cell2mat(array_channels_names(i+1)),"AF"+int2str(record_day+5));
        xlswrite("files/CE.xls",vas_range,cell2mat(array_channels_names(i+1)),"AG"+int2str(record_day+5));
    
        %write in the excel the subject number
        if exist('input_subject')==1
            xlswrite("files/CE.xls",mat2cell(input_subject,1),cell2mat(array_channels_names(i+1)),"A1");
        end
    end

    if ne(isempty(f),1)
        xlswrite("files/"+name_data+"_PSD_CE.xls",mat2cell("Frequency(Hz)",1),"PSD","A2");
        xlswrite("files/"+name_data+"_PSD_CE.xls",f,"PSD","A3");
    end

    if ne(isempty(psd_ce1_total),1)
        xlswrite("files/"+name_data+"_PSD_CE.xls",mat2cell("CE MIN 1",1),"PSD","B1");
        xlswrite("files/"+name_data+"_PSD_CE.xls",array_channels_names(2:end),"PSD","B2");
        xlswrite("files/"+name_data+"_PSD_CE.xls",psd_ce1_total,"PSD","B3");
    end
    if ne(isempty(psd_ce2_total),1)
        xlswrite("files/"+name_data+"_PSD_CE.xls",mat2cell("CE MIN 2",1),"PSD","L1");
        xlswrite("files/"+name_data+"_PSD_CE.xls",array_channels_names(2:end),"PSD","L2");
        xlswrite("files/"+name_data+"_PSD_CE.xls",psd_ce2_total,"PSD","L3");
    end
    if ne(isempty(potencia_ce1),1)
        xlswrite("files/"+name_data+"_PSD_CE.xls",mat2cell("Power",1),"PSD","B265");
        xlswrite("files/"+name_data+"_PSD_CE.xls",potencia_ce1,"PSD","B266");
    end
    if ne(isempty(potencia_ce2),1)
        xlswrite("files/"+name_data+"_PSD_CE.xls",mat2cell("Power",1),"PSD","L265");
        xlswrite("files/"+name_data+"_PSD_CE.xls",potencia_ce2,"PSD","L266");
    end

    %clear all the variables except the ones created at the beginning of
    %the script. We do this to prevent the mixing of information between
    %files
    clearvars -except Fs Fn resolution noverlap nfft window L f Fpass dataf_read data_names eventsf_read events_names
    
end



