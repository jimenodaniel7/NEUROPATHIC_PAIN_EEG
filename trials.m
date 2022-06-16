
 %% Trials WITH ASR

 %this script examine each xlsx file that contain an EEG recording in the
 %directory and extracts all the COMPLETE trials. For each complete trial
 %(the complete trials are the ones in wich the ASR method hadn't take any
 %chunk of signal), it takes 2 seconds to the event (the warning, called "PRE") and three
 %seconds after the event (the trial, called "POST"). Then it stores the information in
 %xlsx, inside a folder called "trials".
 %It is important to remember that, in order to run the fike correctly, you
 %have to fisrt run the bandpass_asr.m and the all_excels_afterasr_woIT.m
 %files. This file has to be run in the same folder where the xls files
 %generated after the bandpass_asr.m file is ran.
 %MATLAB version: 2021b


%We write in a file all the csv data files and in other file the events
%files so we can go extracting data for each file in the following loop
dataf_read=dir(strcat(pwd,'/*-data.xlsx'));
data_names=char(dataf_read.name);
eventsf_read=dir(strcat(pwd,'/*_events.csv'));
events_names=char(eventsf_read.name);





% passband
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
    

    


for d=1:length(data_names(:,1))

    %arrays to store the positions of the left or right arrow events
    positions=[];
    right_arrow_positions=[];
    right_at_positions=[];
    left_arrow_positions=[];
    left_at_positions=[];
    psd_right_1_total=[];
    psd_right_2_total=[];
    psd_left_1_total=[];
    psd_left_2_total=[];
    psd_left_pre_total=[];
    psd_right_pre_total=[];
    psd_trials_right_post=[];
    psd_trials_right_pre=[];
    psd_trials_left_post=[];
    psd_trials_left_pre=[];    
    trials_names_right=[];
    trials_names_left=[];
    trial_fine=[];
    trials_wrong=[];

    kd=strfind(data_names(d,:),".");
    name_data=data_names(d,5:kd-1);
    
    ke=strfind(events_names(d,:),".");
    name_events=events_names(d,5:ke-1);
    count_trials=0;
    count_trials_left=0;
    count_trials_right=0;

    if exist(name_data+"_trials/")~=7
        mkdir(name_data+"_trials/")
    end
    if exist(name_data+"_trials/trials_pre/")~=7
        mkdir(name_data+"_trials/trials_pre/")
    end
    if exist(name_data+"_trials/trials_post/")~=7
        mkdir(name_data+"_trials/trials_post/")
    end

    if exist("C:\\Users\\lacar\\Desktop\\sujeto_10\\asr\\asr_"+name_data+".xlsx")==2
        [data_asr,array_channels_names] = xlsread("C:\\Users\\lacar\\Desktop\\sujeto_10\\asr\\asr_"+name_data+".xlsx");
        bb_data=data_asr;
        bb_data_t=bb_data(:,1); %time vector (in ms)
        bb_data_sintiempo=bb_data(:,2:end); %data vector (only EEG channels)
    end
    
    %events data
    csv_events=readtable(convertCharsToStrings(events_names(end,:)), 'FileType', 'text');
    events=table2array(csv_events(:,1)); %type of the event
    events_latency=table2array(csv_events(:,2));%latency (sample number) of the event
    events_it=events;
    events_it_latency=events_latency;
    
    samples_3_seconds=830; %number of samples in 3 seconds (time between arrow and next warning). The formula followed is 3000000 (time between arrow and warning) /3609 (time between samples)
    
    
    for h=1:length(events_it)
        %if a right arrow event occur, we store the position in right_arrow_positions
        if cell2mat(events_it(h)) == "IMGTASK_RA" && cell2mat(events_it(h+1)) == "IMGTASK_AT" && le((events_it_latency(h+1)-events_it_latency(h)),samples_3_seconds)
           
            position=events_it_latency(h);
            position_r=position;
            position_r_at=events_latency(h+1);
            right_arrow_positions= [right_arrow_positions position_r];
            right_at_positions=[right_at_positions position_r_at];
            positions=[positions position_r position_r_at];
            
        %if a left arrow event occur, we store the position in left_arrow_positions   
        elseif cell2mat(events(h)) == "IMGTASK_LA" && cell2mat(events(h+1)) == "IMGTASK_AT"
            position=events_latency(h);
            position_l=position;
            position_l_at=events_latency(h+1);
            left_arrow_positions= [left_arrow_positions position_l];
            left_at_positions=[left_at_positions position_l_at];
            positions=[positions position_l position_l_at];
          
        end
        %to store the VAS value in a MATLAB variable
        if cell2mat(events(h)) == "VAS"
            vas=events_latency(h);
        end
        %if the event is the PAUSE of the IT part, we store its position
        if cell2mat(events(h)) == "PAUSE" && events_latency(h)>= (length(bb_data_t)*(1/3)) && events_latency(h)<= (length(bb_data_t)*(2/3))
            position_pause=events_latency(h);
            positions=[positions position_pause];
        end
    end
    
    
    
    %this loop writes the data for every channel
    
    samples_3_seconds=830; %number of samples in 3 seconds (time between arrow and next warning). The formula followed is 3000000 (time between arrow and warning) /3609 (time between samples)
    samples_2_seconds=550; %number of samples in 3 seconds (time between arrow and next warning). The formula followed is 2000000 (time between arrow and warning) /3609 (time between samples)
    for r= 1:length(positions)

        if (isempty(find(right_arrow_positions==positions(r)))==0)
            count_trials=count_trials+1;
            count_trials_right=count_trials_right+1;
            bb_data_sintiempo_r_post=bb_data_sintiempo(positions(r):positions(r)+samples_3_seconds,:);
            bb_data_sintiempo_r_pre=bb_data_sintiempo(positions(r)-samples_2_seconds:positions(r),:);
            [psd_right_post,f_right_post]=pwelch(bb_data_sintiempo_r_post,window,[],nfft,Fs);
            [psd_right_pre,f_right_pre]=pwelch(bb_data_sintiempo_r_pre,window,[],nfft,Fs);
            %for every position store in the array of positions, we extract the
            %data between the event and the next warning (approximately 3 seconds, which correspond with a little more than 830 samples)
            
            %PSD RIGHT PRE%
            figure
            for i=1:(length(array_channels_names)-1)
                subplot(1,3,1)
                if array_channels_names(i+1)=="FC3" || array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="CP3" 
                    plot (f_right_pre (5:80), psd_right_pre(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD left channels ")
                    axis([0 40 0 15])
                    hold on
                end
                subplot(1,3,2)
                if array_channels_names(i+1)=="FCz" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="CPZ" 
                    plot (f_right_pre (5:80), psd_right_pre(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD central channels ")
                    axis([0 40 0 15])
                    hold on
    
                end
                subplot(1,3,3)
                if array_channels_names(i+1)=="FC4" || array_channels_names(i+1)=="C4" || array_channels_names(i+1)=="CP4" 
                    plot (f_right_pre (5:80), psd_right_pre(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD right channels ")
                    axis([0 40 0 15])
                    hold on
                end
                sgtitle(name_data+" Total trials "+count_trials+ " RIGHT ARROW PRE trial "+count_trials_right)
                
                saveas(gcf,name_data+"_trials/trials_pre/"+name_data+"_"+count_trials+"_RIGHT_ARROW_PRE_"+count_trials_right+".jpg")
                if array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="C4" 
                    psd_trials_right_pre=[psd_trials_right_pre psd_right_pre(:,i)];

                end
            end

            close


            %PSD RIGHT POST%
            figure
            for i=1:(length(array_channels_names)-1)
                subplot(1,3,1)
                if array_channels_names(i+1)=="FC3" || array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="CP3" 
                    plot (f_right_post (5:80), psd_right_post(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD left channels ")
                    axis([0 40 0 15])
                    hold on
                end
                subplot(1,3,2)
                if array_channels_names(i+1)=="FCz" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="CPZ" 
                    plot (f_right_post (5:80), psd_right_post(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD central channels ")
                    axis([0 40 0 15])
                    hold on
    
                end
                subplot(1,3,3)
                if array_channels_names(i+1)=="FC4" || array_channels_names(i+1)=="C4" || array_channels_names(i+1)=="CP4" 
                    plot (f_right_post (5:80), psd_right_post(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD right channels ")
                    axis([0 40 0 15])
                    hold on
                end
                sgtitle(name_data+" Total trials "+count_trials+ " RIGHT ARROW POST trial "+count_trials_right)
                
                saveas(gcf,name_data+"_trials/trials_post/"+name_data+"_"+count_trials+"_RIGHT_ARROW_POST_"+count_trials_right+".jpg")
                if array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="C4" 
                    psd_trials_right_post=[psd_trials_right_post psd_right_post(:,i)];
                    name_right=array_channels_names(i+1)+"T"+num2str(count_trials_right);
                    trials_names_right=[trials_names_right name_right];
                end
            end

            close
         %we store a line of 0s in the place in the middle of the array,
         %right where the PAUSE will be.
        elseif exist('position_pause') && positions(r)==position_pause
            psd_trials_right_pre=[psd_trials_right_pre zeros(257,1)];
            psd_trials_right_post=[psd_trials_right_post zeros(257,1)];
            psd_trials_left_pre=[psd_trials_left_pre zeros(257,1)];
            psd_trials_left_post=[psd_trials_left_post zeros(257,1)];
            trials_names_left=[trials_names_left "PAUSE"];
            trials_names_right=[trials_names_right "PAUSE"];
            
        elseif (isempty(find(left_arrow_positions==positions(r)))==0)
            count_trials=count_trials+1;
            count_trials_left=count_trials_left+1;
            bb_data_sintiempo_l_post=bb_data_sintiempo(positions(r):positions(r)+samples_3_seconds,:);
            bb_data_sintiempo_l_pre=bb_data_sintiempo(positions(r)-samples_2_seconds:positions(r),:);
            [psd_left_post,f_left_post]=pwelch(bb_data_sintiempo_l_post,window,[],nfft,Fs);
            [psd_left_pre,f_left_pre]=pwelch(bb_data_sintiempo_l_pre,window,[],nfft,Fs);
            
            %PSD LEFT PRE%
            figure
            for i=1:(length(array_channels_names)-1)
                subplot(1,3,1)
                if array_channels_names(i+1)=="FC3" || array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="CP3" 
                    plot (f_left_pre (5:80), psd_left_pre(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD left channels ")
                    axis([0 40 0 15])
                    hold on
                end
                subplot(1,3,2)
                if array_channels_names(i+1)=="FCz" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="CPZ" 
                    plot (f_left_pre (5:80), psd_left_pre(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD central channels ")
                    axis([0 40 0 15])
                    hold on
                end
                subplot(1,3,3)
                if array_channels_names(i+1)=="FC4" || array_channels_names(i+1)=="C4" || array_channels_names(i+1)=="CP4" 
                    plot (f_left_pre (5:80), psd_left_pre(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD right channels ")
                    axis([0 40 0 15])
                    hold on
                end
                sgtitle(name_data+" Total trials "+count_trials+ " LEFT ARROW PRE trial "+count_trials_left)
                legend show
                saveas(gcf,name_data+"_trials/trials_pre/"+name_data+"_"+count_trials+"_LEFT_ARROW_PRE_"+count_trials_left+".jpg")
                if array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="C4" 
                    psd_trials_left_pre=[psd_trials_left_pre psd_left_pre(:,i)];

                end
            end
            close

            
            %PSD LEFT POST%
            figure
            for i=1:(length(array_channels_names)-1)
                subplot(1,3,1)
                if array_channels_names(i+1)=="FC3" || array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="CP3" 
                    plot (f_left_post (5:80), psd_left_post(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD left channels ")
                    axis([0 40 0 15])
                    hold on
                end
                subplot(1,3,2)
                if array_channels_names(i+1)=="FCz" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="CPZ" 
                    plot (f_left_post (5:80), psd_left_post(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD central channels ")
                    axis([0 40 0 15])
                    hold on
                end
                subplot(1,3,3)
                if array_channels_names(i+1)=="FC4" || array_channels_names(i+1)=="C4" || array_channels_names(i+1)=="CP4" 
                    plot (f_left_post (5:80), psd_left_post(5:80,i),'DisplayName',cell2mat(array_channels_names(i+1)));
                    legend show
                    title("PSD right channels ")
                    axis([0 40 0 15])
                    hold on
                end
                sgtitle(name_data+" Total trials "+count_trials+ " LEFT ARROW POST trial "+count_trials_left)
                legend show
                saveas(gcf,name_data+"_trials/trials_post/"+name_data+"_"+count_trials+"_LEFT_ARROW_POST_"+count_trials_left+".jpg")
                if array_channels_names(i+1)=="C3" || array_channels_names(i+1)=="Cz" || array_channels_names(i+1)=="C4" 
                    psd_trials_left_post=[psd_trials_left_post psd_left_post(:,i)];
                    name_left=array_channels_names(i+1)+"T"+num2str(count_trials_left);
                    trials_names_left=[trials_names_left name_left];
                end
            end
            close    

        end
    end
    close all
    
    %all the information is stored in xls files. At the end, we store the
    %PSD of all trials (PRE and POST) in a file.
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",psd_trials_right_pre,"pre_trials_PSD","A2");
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",trials_names_right,"pre_trials_PSD","A1");
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",psd_trials_right_post,"post_trials_PSD","A2");
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",trials_names_right,"post_trials_PSD","A1");
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",psd_trials_left_pre,"pre_trials_PSD","A2");
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",trials_names_left,"pre_trials_PSD","A1");
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",psd_trials_left_post,"post_trials_PSD","A2");
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",trials_names_left,"post_trials_PSD","A1");
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",mat2cell("VAS",1),"pre_trials_PSD","A"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",vas,"pre_trials_PSD","B"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",mat2cell("VAS",1),"post_trials_PSD","A"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_RIGHT.xls",vas,"post_trials_PSD","B"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",mat2cell("VAS",1),"pre_trials_PSD","A"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",vas,"pre_trials_PSD","B"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",mat2cell("VAS",1),"post_trials_PSD","A"+int2str(length(psd_trials_right_post(:,1))+5));
    xlswrite(name_data+"_trials/"+name_data+"_trials_LEFT.xls",vas,"post_trials_PSD","B"+int2str(length(psd_trials_right_post(:,1))+5));
end
