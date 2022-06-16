%% Script that load filter the EEG signals

% This script load all the csv files that the EEG device
% (Bitbrain Hero with 9 channels) provide to filter them using various methods. 
% When ran, it generates a
% series of folders with the figures of the PSD (Power Spectrum Density) of
% the raw signal (w/o any filter) and some xls files with the signal filtered
% passband. It has to be ran in the same folder as the files from the EEG
% device. It is important to mention that the script uses an Add-ON EEGLAB,
% so it has to be previously install in your computer in order to use the
% script correctly. In EEGLAB, it is also necessary to install a Plug-In to
% use the ASR method. Please notice that you have tu run this script before
% the rest of them.
%MATLAB version: 2021b
%EEGLAB version: 2021.1
%Plug-In: Clean RaWData and ASR Method

route="C:\\Users\\lacar\\Desktop\\sujeto_10\\";
% Frequencies of the bandpass filter
Fpass = [0.5 40]; 

% Sampling Frequency
Fs = 256;   

% Nyquist Frequency
Fn = Fs/2; 
resolution=0.5;
noverlap=0;
nfft=round(Fs/resolution);
window=hanning(nfft);
L = 1200;
    
%We write in a struct all the csv data files and in other struct the events
%files so we can go extracting data for each file in the following loop
dataf_read=dir(strcat(route+'*-data.csv'));
data_names=char(dataf_read.name);
eventsf_read=dir(strcat(route+'*-events.csv'));
events_names=char(eventsf_read.name);

%Folder to store the raw data matrix
if exist(route+'raw_data')~=7
   mkdir(route+'raw_data')
end



%Folder to store data after using a bandpass filter (0.5-40 Hz)
if exist(route+'bandpass')~=7
    mkdir(route+'bandpass')
end

%Folder to store de data after using ASR
if exist(route+'asr')~=7
    mkdir(route+'asr')
end

%loop to filter each signal
for d=1:length(data_names(:,1))

    %to use the name of the files
    kd=strfind(data_names(d,:),".");
    name_data=data_names(d,1:kd-1);
    
    ke=strfind(events_names(d,:),".");
    name_events=events_names(d,1:ke-1);

    bitbrain_o=readtable(convertCharsToStrings(route+data_names(d,:)) ,'FileType', 'text');
    bitbrain=table2array(bitbrain_o);
    bitbrain_data_t=bitbrain(:,1); %time vector
    bitbrain_data_t_corrected=bitbrain_data_t-bitbrain_data_t(1,1); %time vector corrected so it match the one used in EEGLAB (that starts in 0s)


    
    bitbrain_data=bitbrain(:,:); %all channels

    %to save the raw data matrix in the folder created before
    csvwrite(route+"raw_data\\"+name_data+"_raw.csv", bitbrain_data);
    
    
    %we read the events csv and we correct it, so it match the times vector
    %used in EEGLAB
    csv_events=readtable(convertCharsToStrings(route+events_names(d,:)), 'FileType', 'text');
    events=table2array(csv_events(:,1));
    events_corrected=events-bitbrain_data(1,1);
  
   

    bb_data_sintiempo=bandpass(bitbrain_data(:,2:end),  Fpass, Fs);
    bb_data=[bitbrain_data_t_corrected,bb_data_sintiempo];

    %positions of events in data array. This formula gives the position of the
    %closest point in the time vector to the trigger event.
    
    %position of the first oe event
    [min_difference_oe1, position_oe1] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(4)));
    
    %position of the ce event
    [min_difference_ce, position_ce] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(5)));
    [min_difference_p1, position_p1] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(6))); %first pause
    
    %position of start of first task event
    [min_difference_t1, position_t1] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(7)));
    [min_difference_p2, position_p2] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(127))); %second pause
    
    %position of start of second task event
    [min_difference_t2, position_t2] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(128)));
    [min_difference_p3, position_p3] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(247))); %third pause
    
    %position of the second oe event
    [min_difference_oe3, position_oe3] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(248)));
    [min_difference_end, position_end] = min(abs(bitbrain_data_t_corrected(:,1)-events_corrected(249))); %end


    %we calculate extra positions to divide the events
    position_oe2=(position_oe1 + position_ce) /2;
    position_oe4=(position_oe3 + position_end) /2;
    position_ce2=(position_ce + position_p1) /2;

    samples_60_seconds=15360; %number of samples in 60 seconds. The formula followed is 60s * Fs

    %to create folders to store the figures of psd before and after ASR
    if exist(route+name_data+"_figures")~=7
        mkdir(route+name_data+"_figures")
        mkdir(route+name_data+"_figures\\figures_asr")
        mkdir(route+name_data+"_figures\\figures_sinasr")
    end

    %%PSD FIGURES (before ASR)%%
    %% 1OE 
    [psd,f]=pwelch(bb_data(position_oe1:position_oe2,2:end),window,[],nfft,Fs);
    subplot(1,3,1)
    plot (f (5:80), psd(5:80,[1,4,7]));
    title("PSD left channels")
    legend("FC3", "C3", "CP3")
    axis([0 29 0 15]) %we focus in the frequencies from 0 to 29 Hz (alpha, beta and theta)
    subplot(1,3,2)
    plot (f (5:80), psd(5:80,[2,5,8]));
    title("PSD central channels")
    axis([0 29 0 15])
    legend("FCz","Cz", "CPz")
    subplot(1,3,3)
    plot (f (5:80), psd(5:80,[3,6,9]));
    title("PSD right channels")
    axis([0 29 0 15])
    legend("FC4", "C4", "CP4")
    sgtitle(name_data+" 10E SIN ASR")
    saveas(gcf,route+name_data+"_figures\\figures_sinasr\\"+name_data+"_PSD_1OE_sinasr.jpg")
    
   
    %% 2OE 
    [psd,f]=pwelch(bb_data(position_oe2:position_ce,2:end),window,[],nfft,Fs);
    subplot(1,3,1)
    plot (f (5:80), psd(5:80,[1,4,7]));
    title("PSD left channels")
    axis([0 29 0 15])
    legend("FC3", "C3", "CP3")
    subplot(1,3,2)
    plot (f (5:80), psd(5:80,[2,5,8]));
    title("PSD central channels")
    axis([0 29 0 15])
    legend("FCz","Cz", "CPz")
    subplot(1,3,3)
    plot (f (5:80), psd(5:80,[3,6,9]));
    title("PSD right channels")
    axis([0 29 0 15])
    legend("FC4", "C4", "CP4")
    sgtitle(name_data+" 20E SIN ASR")
    saveas(gcf,route+name_data+"_figures\\figures_sinasr\\"+name_data+"_PSD_2OE_sinars.jpg")

   %% 1CE 
    [psd,f]=pwelch(bb_data(position_ce:position_ce2,2:end),window,[],nfft,Fs);
    subplot(1,3,1)
    plot (f (5:80), psd(5:80,[1,4,7]));
    title("PSD left channels")
    axis([0 29 0 15])
    legend("FC3", "C3", "CP3")
    subplot(1,3,2)
    plot (f (5:80), psd(5:80,[2,5,8]));
    title("PSD central channels")
    axis([0 29 0 15])
    legend("FCz","Cz", "CPz")
    subplot(1,3,3)
    plot (f (5:80), psd(5:80,[3,6,9]));
    title("PSD right channels")
    axis([0 29 0 15])
    legend("FC4", "C4", "CP4")
    sgtitle(name_data+" 1CE SIN ASR")
    saveas(gcf,route+name_data+"_figures\\figures_sinasr\\"+name_data+"_PSD_1CE_sinars.jpg")
   %% 2CE 
    [psd,f]=pwelch(bb_data(position_ce2:position_p1,2:end),window,[],nfft,Fs);
    subplot(1,3,1)
    plot (f (5:80), psd(5:80,[1,4,7]));
    title("PSD left channels")
    axis([0 29 0 15])
    legend("FC3", "C3", "CP3")
    subplot(1,3,2)
    plot (f (5:80), psd(5:80,[2,5,8]));
    title("PSD central channels")
    axis([0 29 0 15])
    legend("FCz","Cz", "CPz")
    subplot(1,3,3)
    plot (f (5:80), psd(5:80,[3,6,9]));
    title("PSD right channels")
    axis([0 29 0 15])
    legend("FC4", "C4", "CP4")
    sgtitle(name_data+" 2CE SIN ASR")
    saveas(gcf,route+name_data+"_figures\\figures_sinasr\\"+name_data+"_PSD_2CE_sinars.jpg")
    %% 3OE 
    [psd,f]=pwelch(bb_data(position_oe3:position_oe4,2:end),window,[],nfft,Fs);
    subplot(1,3,1)
    plot (f (5:80), psd(5:80,[1,4,7]));
    title("PSD left channels")
    axis([0 29 0 15])
    legend("FC3", "C3", "CP3")
    subplot(1,3,2)
    plot (f (5:80), psd(5:80,[2,5,8]));
    title("PSD central channels")
    axis([0 29 0 15])
    legend("FCz","Cz", "CPz")
    subplot(1,3,3)
    plot (f (5:80), psd(5:80,[3,6,9]));
    title("PSD right channels")
    axis([0 29 0 15])
    legend("FC4", "C4", "CP4")
    sgtitle(name_data+" 30E SIN ASR")
    saveas(gcf,route+name_data+"_figures\\figures_sinasr\\"+name_data+"_PSD_30E_sinars.jpg")
   
    %% 4OE 
    [psd,f]=pwelch(bb_data(position_oe4:(position_oe4+samples_60_seconds),2:end),window,[],nfft,Fs);
    subplot(1,3,1)
    plot (f (5:80), psd(5:80,[1,4,7]));
    title("PSD left channels")
    axis([0 29 0 15])
    legend("FC3", "C3", "CP3")
    subplot(1,3,2)
    plot (f (5:80), psd(5:80,[2,5,8]));
    title("PSD central channels")
    axis([0 29 0 15])
    legend("FCz","Cz", "CPz")
    subplot(1,3,3)
    plot (f (5:80), psd(5:80,[3,6,9]));
    title("PSD right channels")
    axis([0 29 0 15])
    legend("FC4", "C4", "CP4")
    sgtitle(name_data+" 40E SIN ASR")
    saveas(gcf,route+name_data+"_figures\\figures_sinasr\\"+name_data+"_PSD_40E_sinars.jpg")
    close all

    %channels names store in a str array
    array_channels_names=["FC3", "FCZ", "FC4", "C3", "Cz", "C4", "CP3", "CPz", "CP4"]; %Use in the for loop to create the different files

    
    %We need to create a table with all the events in the script, along
    %with their duration and their latency. With all that information, we
    %will be able to import those events in EEGLAB, so we don't lose any
    %information of any event.

    % Events table to import in EEGLAB
    type=table2array(csv_events(3:end,2)); %name of the event
    events_eeglab=events_corrected(3:end);
    latency=zeros(length(type),1); %latency of the events (sample number)
    left_arrow_positions=[];
    right_arrow_positions=[];
    duration=zeros(length(type),1); %duration of the events in number of samples
    for h=1:length(type)
        if cell2mat(type(h)) == "OE" 
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            duration(h)=120*256; %samples in 2 minutes
            latency(h)=position;
        elseif cell2mat(type(h)) == "CE"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            duration(h)=120*256; %samples in 2 minutes
            latency(h)=position; 
        elseif cell2mat(type(h)) == "IMGTASK_START"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            duration(h)=5*256; %samples in 5 seconds
            latency(h)=position;
        elseif cell2mat(type(h)) == "IMGTASK_RA"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            position_r=position+1;
            right_arrow_positions= [right_arrow_positions position_r];
            duration(h)=3*256; %samples in 3 seconds
            latency(h)=position;
        elseif cell2mat(type(h)) == "IMGTASK_LA"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            position_l=position+1;
            left_arrow_positions= [left_arrow_positions position_l];
            duration(h)=3*256; %samples in 3 seconds
            latency(h)=position;
        elseif cell2mat(type(h)) ==  "IMGTASK_AT"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            duration(h)=2*256; %samples in 2 seconds
            latency(h)=position;
        elseif cell2mat(type(h)) ==  "PAUSE"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            duration(h)=1*256; %samples in 1 second (just to mark an event there)
            latency(h)=position;
        elseif cell2mat(type(h)) ==  "END"
            [min_difference, position] = min(abs(bitbrain_data_t_corrected(:,1)-events_eeglab(h,1)));
            duration(h)=1*256; %samples in 1 second (just to mark an event there)
            latency(h)=position;
        end
    end
    eventos=table(type,latency,duration); %table of events
    eventos_size=size(eventos);


    %we start EEGLAB to use the ASR method
    EEG.etc.eeglabvers = '2021.1'; % this tracks which version of EEGLAB is being used, you may ignore it
    EEG = pop_importdata('dataformat','array','nbchan',0,'data',bb_data_sintiempo','srate',256,'pnts',0,'xmin',0,'chanlocs','C:\\Users\\lacar\\Desktop\\sujeto_10\\Painapp.sfp'); %import data without time vector so we can import channel information also

    EEG.setname=name_data; %name the dataset
    EEG = eeg_checkset( EEG );
    EEG=pop_chanedit(EEG, 'lookup','C:\\Users\\lacar\\Desktop\\sujeto_10\\Painapp.sfp');
    EEG = eeg_checkset( EEG );

    %import the events table in EEGLAB
    for e=1:eventos_size(1)
         n_events=length(EEG.event);
         EEG.event(n_events+1).type=cell2mat(table2cell(eventos(e,1)));
         EEG.event(n_events+1).latency=cell2mat(table2cell(eventos(e,2)));
         EEG.event(n_events+1).duration=cell2mat(table2cell(eventos(e,3)));
    end

    %check for consistency and reorder the events chronologically
    EEG=eeg_checkset(EEG,'eventconsistency');
    %save the dataset before using ASR (only with bp filter)
    EEG = pop_saveset( EEG, 'filename',name_data,'filepath','C:\\Users\\lacar\\Desktop\\sujeto_10\\bandpass\\');
    tiempo_eeglab_sinars=EEG.times;

    %asr method
    EEG = pop_clean_rawdata(EEG, 'FlatlineCriterion',5,'ChannelCriterion',1,'LineNoiseCriterion',4,'Highpass','off','BurstCriterion',10,'WindowCriterion',0.25,'BurstRejection','on','Distance','Euclidian','WindowCriterionTolerances',[-Inf 7] );
    EEG.setname=name_data+"-asr";
    EEG = eeg_checkset( EEG);
    array_channels={EEG.chanlocs.labels};
    %show the resulting PSD (after using ASR)
    pop_spectopo(EEG, 1, [0      913723.4689], 'EEG' , 'percent', 50, 'freqrange',[2 40],'electrodes','off');
    legend(array_channels)
    title("PSD after ASR")
    saveas(gcf,route+name_data+"_figures\\figures_asr\\"+name_data+"_PSD_ARS.jpg")


    name_asr=strcat('asr_',name_data);




    

    %save dataset and csv after ASR
    EEG = pop_saveset( EEG, 'filename',name_asr,'filepath','C:\\Users\\lacar\\Desktop\\sujeto_10\\asr\\');
    bb_data_asr=[EEG.times; EEG.data]';
    xlswrite(route+"asr\\"+name_asr+".xlsx", bb_data_asr,1,"A2");
    
    letters=["A" "B" "C" "D" "E" "F" "G" "H" "I" "J" "K"];
    for i=1:length(array_channels)
        xlswrite(route+"asr\\"+name_asr+".xlsx", array_channels(i),1,letters(i+1));
    end
    xlswrite(route+"asr\\"+name_asr+".xlsx", mat2cell("Time",1),1,"A1");


    %save events in csv file
    nonzero_values=find(EEG.etc.clean_sample_mask);


    tipo={EEG.event.type,  "VAS"}';
    tipo_array=table2array(cell2table(tipo));
    size_tipo=size(tipo);
    latencia=[table2array(table(EEG.event.latency)),str2double(table2array(csv_events(2,3)))]';
    duracion=[table2array(table(EEG.event.duration)),"-"]';

    %here, we first make sure that the "important events" in the csv file
    %are still there or the ASR method had removed them. If not, we create a new event with the same name and
    %store it in the place it should be.

    %check or store the fisrt OE event
    if isempty(find(tipo_array(1:round(length(tipo_array)/2),1)=="OE")) && nonzero_values(1,1)<(45*256)
        [min_difference_oe, p_oe1] = min(abs(nonzero_values(1,:)-position_oe1));
        if ne(p_oe1,1)
            position_oe1_asr=p_oe1;
        else
            position_oe1_asr= 2*256;
        end
        for i=2:length(latencia)
            if latencia(i-1) < position_oe1_asr && latencia(i) > position_oe1_asr
                la=[latencia(1:i-1); position_oe1_asr; latencia(i:end)];
                du=[duracion(1:i-1); 30720; duracion(i:end)];
                ti=[tipo_array(1:i-1); "OE"; tipo_array(i:end)];
                break
            end
        end
        latencia=la;
        duracion=du;
        tipo_array=ti;
        tipo=table2cell(array2table(tipo_array));
    end
    %check or store the CE event
    if isempty(find(tipo_array(1:end-1,1)=="CE")) && nonzero_values(1,1)<(165*256)
        [min_difference_ce, position_ce_asr] = min(abs(nonzero_values(1,:)-position_ce));
        for i=2:length(latencia)
            if latencia(i-1) < position_ce_asr && latencia(i) > position_ce_asr
                la=[latencia(1:i-1); position_ce_asr; latencia(i:end)];
                du=[duracion(1:i-1); 30720; duracion(i:end)];
                ti=[tipo_array(1:i-1); "CE"; tipo_array(i:end)];
                break
            end
        end
        latencia=la;
        duracion=du;
        tipo_array=ti;
        tipo=table2cell(array2table(tipo_array));
    end
    %check or store the fisrt PAUSE event
    if isempty(find(tipo_array(1:round(length(tipo_array)/2),1)=="PAUSE"))
        [min_difference_p1, position_p1_asr] = min(abs(nonzero_values(1,:)-position_p1));
        for i=2:length(latencia)
            if latencia(i-1) < position_p1_asr && latencia(i) > position_p1_asr
                la=[latencia(1:i-1); position_p1_asr; latencia(i:end)];
                du=[duracion(1:i-1); 1; duracion(i:end)];
                ti=[tipo_array(1:i-1); "PAUSE"; tipo_array(i:end)];
                break
            end
        end
        latencia=la;
        duracion=du;
        tipo_array=ti;
        tipo=table2cell(array2table(tipo_array));
    end
    %check or store the SECOND OE event
    if isempty(find(tipo_array(round(length(tipo_array)/2):end-1,1)=="OE"))
        [min_difference_oe3, position_oe3_asr] = min(abs(nonzero_values(1,:)-position_oe3));
        for i=2:length(latencia)
            if latencia(i-1) < position_oe3_asr && latencia(i) > position_oe3_asr
                la=[latencia(1:i-1); position_oe3_asr; latencia(i:end)];
                du=[duracion(1:i-1); 30720; duracion(i:end)];
                ti=[tipo_array(1:i-1); "OE"; tipo_array(i:end)];
                break
            end
        end
        latencia=la;
        duracion=du;
        tipo_array=ti;
        tipo=table2cell(array2table(tipo_array));
    end
    
    %we write the events in a table and export them to a csv file
    eventos_asr=table(tipo,latencia,duracion);
    writetable(eventos_asr,route+"asr\\"+name_asr+"_events.csv");

end
