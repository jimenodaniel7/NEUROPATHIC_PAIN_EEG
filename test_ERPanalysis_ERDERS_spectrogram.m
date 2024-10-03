%START EEGLAB
[ALLEEG EEG CURRENTSET ALLCOM] = eeglab; 
eeglab;


% This file is a demo for testing the ERD/ERS analysis functions
% This function is produced by Esmaeil Seraj (esmaeil.seraj09@gmail.com)
% 
% Dependencies: Functions "bdf2mat_main.mat", "emg_onset.mat", 
%               "trigger_avg_erp.mat", "BPFilter5.mat", "BaseLine2.mat"
%               and "sig_trend.mat" (provided by the same author). Also 
%               "eeg_read_bdf.mat" (provided by Gleb Tcheslavski
%               (gleb@vt.edu))
% 
% ***NOTE***: Confidential Content, please DO NOT modify or redistribute
%             without permision of the producer. Property of GT-Bionics Lab
%             ECE, Georgia Tech, Atlanta, Georgia, U.S.
%

close all
clear
clc

% Specify the directory to search
directoryPath = 'C:\Users\jimen\OneDrive\Escritorio\HNPT\EEG_PAINAPP\SUJETOS'; % Change this to your target directory

% Get all items in the directory
allItems = dir(directoryPath);

% Filter out items that are not directories and exclude '.' and '..'
folders = allItems([allItems.isdir]);  % Only directories
folders = folders(~ismember({folders.name}, {'.', '..'}));  % Exclude '.' and '..'





%% ERP time-course estimation through triggered-average of EEG ensembles
duration = 2;    % required signal duration after movement onset in seconds
freq_band = 'beta';                  % available options: 'alpha' or 'beta'
emg={};
ref_per = [-1, 0]; % double vector in [-a, -b] form where -a and -b are the edges of reference segment
cof_intv = 3;                             % confidence interval coefficient
fs=256;
%1 FC4,2 C3,3 Cz,4 C4,5 CP3,6 CP4
right=1;
left=2;


% Generate array with subjects
% folderPath='C:\Users\jimen\OneDrive\Escritorio\HNPT\EEG_PAINAPP\Sub_3';
filePath = 'C:\Users\jimen\OneDrive\Escritorio\HNPT\EEG_PAINAPP\LIST_IT.xlsx';

for s=1:length(folders)
    folderPath=fullfile(folders(s).folder,folders(s).name);
    a=dir(strcat(folderPath,'/*.set'));
    d=char(a.name);
    
    
    
    subject_data = select_good_channels(filePath);
    
    % Check if the folder exists
    output_figures = fullfile(folderPath,'Figures');
    if ~exist(output_figures , 'dir')
        % If the folder doesn't exist, create it
        mkdir(output_figures);
    end
    
    
    
    % Generate some variables
    trigger_time_sec_right=1;
    trigger_time_sec_left=1;
    trigger_time_sec_rightt=1;
    columns=char('A':'Z');
    
    % Arrays to store the data
    left_combined_PSD_C3=[];
    left_combined_PSD_Cz=[];
    left_combined_PSD_C4=[];
    left_combined_PSD_CP3=[];
    left_combined_PSD_CP4=[];
    left_combined_PSD_CPz=[];
    left_combined_PSD_FC3=[];
    left_combined_PSD_FC4=[];
    left_combined_PSD_FCz=[];
    
    right_combined_PSD_C3=[];
    right_combined_PSD_Cz=[];
    right_combined_PSD_C4=[];
    right_combined_PSD_CP3=[];
    right_combined_PSD_CP4=[];
    right_combined_PSD_CPz=[];
    right_combined_PSD_FC3=[];
    right_combined_PSD_FC4=[];
    right_combined_PSD_FCz=[];
    

    %%
    % Compute ERP and ERD-ERS ratio for all experiments of the subject
    for i=1:length(d(:,1)) 
        erd_ers_right={};
        erd_ers_left={};
        [PATHSTR,NAME,EXT]=fileparts(d(i,:));
    
    
        % TO TAKE JUST THE GOOD CHANNELS IN THE EXCEL
        n=split(NAME,'_');
    
        subject= folders(s).name;
    
        files = {subject_data.(subject).file};
        % Find the index of the file inside this subject's file list
        for j = 1:length(files)
            if strcmp(files{j}, n(2))
                file_index = j;  % Store the index of the file
                break;  % Exit the loop once the file is found
            end
        end
        good_channels=subject_data.(subject)(file_index).good_channels; % This are the good channels for that file
        
        % If there are no good channels, pass to the next file
        if isempty(good_channels)
            continue
        end
    
        % Create filename array to introduce in Excel
        if i==1
            filename = [{NAME}, num2cell(NaN(1,8))];
        else
            filename = [filename,{NAME}, num2cell(NaN(1,8))];
        end
    
        % Compute ERP of the subject
        [EEG, ERP, ERP_TFFT, ERP_EFFT]=erp_bin(strcat(NAME,EXT),folderPath, pwd);
        good_channels = good_channels(ismember(good_channels,{EEG.chanlocs.labels}));
        
        
        letter=0;
    
        for c=1:length(ERP.chanlocs)
            
            % Select name of channel
            canal = ERP.chanlocs(c).labels;
    
            % IF the channel store in the variable canal is among the good
            % channels
            if ismember(canal, good_channels)
                
                %Separate ERP of left and right trials
                erp_right=ERP.bindata(:,:,right);
                erp_left=ERP.bindata(:,:,left);
                erp_rightt=ERP.bindata(:,:,left);
                erp_right_c=erp_right(c,:);
                erp_left_c = erp_left(c,:);
                
                
                erp_rightt=erp_rightt(c,:);
                time_vec_rightt=EEG.times;
                time_vec_left=EEG.times;
                time_vec_right=EEG.times;
                
                % Smooth ERP power
                smoothed_erp_right = moving_average(erp_right(c,:), 5);
                smoothed_erp_left = moving_average(erp_left(c,:), 5);
                name_fig=strcat(NAME,'.fig');
                
                savefig(fullfile(output_figures,name_fig))
                title={strcat('right_',canal);strcat('left_',canal);};
                letter=letter+1;
        
                %We write the ERP in an Excel file
                xlswrite(fullfile(folderPath,'ERPs.xls'), title', NAME, strcat (columns(letter), '1'));
                
                xlswrite(fullfile(folderPath,'ERPs.xls'), erp_right_c', NAME, strcat (columns(letter), '2'));%right canal 
                letter=letter+1;
        
                xlswrite(fullfile(folderPath,'ERPs.xls'), erp_left_c', NAME, strcat (columns(letter), '2'));%left canal 
                
                % Extract baseline and post-stimulus data from the averaged data
                baseline_data_right = smoothed_erp_right(1:256);
                post_stimulus_data_right = smoothed_erp_right(256:end);
                baseline_data_left  = smoothed_erp_left(1:256);
                post_stimulus_data_left  = smoothed_erp_left(256:end);
                
                
                
                % Compute power spectral density for baseline and post-stimulus periods
                % With this formula, ERD are negative values, ERS are positive values
                erd_ers_right_channel = ((smoothed_erp_right-mean(baseline_data_right))./mean(baseline_data_right)).*100;
                erd_ers_left_channel = ((smoothed_erp_right-mean(baseline_data_left)  )./mean(baseline_data_left) ).*100;
                
                % Store the results for the current channel
                if mean(baseline_data_right)==0
                erd_ers_right{end+1}= zeros(size(erd_ers_right_channel));
                else
                erd_ers_right{end+1} = erd_ers_right_channel;
                end
                if mean(baseline_data_left)  ==0
                erd_ers_left{end+1}= zeros(size(erd_ers_left_channel));
                else
                erd_ers_left{end+1} = erd_ers_left_channel;
                end
            else
                continue
            end 
        end
    
        [psd_left,freq] = plot_spectrogram(erd_ers_left, fs, [8,30],'Left Hand Motor Imagery',good_channels);
        
        name_fig_psd_left=strcat(NAME,'_ERD_ERS_left.png');
    
        saveas(gcf, fullfile(output_figures, name_fig_psd_left));  % Save as PNG file
        
        [psd_right,freq] = plot_spectrogram(erd_ers_right, fs, [8,30],'Right Hand Motor Imagery',good_channels);
            
        name_fig_psd_right=strcat(NAME,'_ERD_ERS_right.png');
        saveas(gcf, fullfile(output_figures, name_fig_psd_right));  % Save as PNG file
       
        
        
        for c = 1:length(good_channels)
            if(strcmp((EEG.chanlocs(c).labels),'C3'))
                canal='C3';
                if ismember(canal, good_channels)
                    left_combined_PSD_C3= [left_combined_PSD_C3, psd_left{c}];
                    right_combined_PSD_C3= [right_combined_PSD_C3, psd_right{c}];
                    right_combined_PSD_C3= [right_combined_PSD_C3, NaN(length(psd_right{c}),1)];
                    left_combined_PSD_C3= [left_combined_PSD_C3, NaN(length(psd_left{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'C4'))
                canal='C4';
                if ismember(canal, good_channels)
                    left_combined_PSD_C4= [left_combined_PSD_C4, psd_left{c}];
                    right_combined_PSD_C4= [right_combined_PSD_C4, psd_right{c}];
                    left_combined_PSD_C4= [left_combined_PSD_C4, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_C4= [right_combined_PSD_C4, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'Cz'))
                canal='Cz';
                if ismember(canal, good_channels)
                    left_combined_PSD_Cz= [left_combined_PSD_Cz, psd_left{c}];
                    right_combined_PSD_Cz= [right_combined_PSD_Cz, psd_right{c}];
                    left_combined_PSD_Cz= [left_combined_PSD_Cz, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_Cz= [right_combined_PSD_Cz, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'CP3'))
                canal='CP3';
                if ismember(canal, good_channels)
                    left_combined_PSD_CP3= [left_combined_PSD_CP3, psd_left{c}];
                    right_combined_PSD_CP3= [right_combined_PSD_CP3, psd_right{c}];
                    left_combined_PSD_CP3= [left_combined_PSD_CP3, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_CP3= [right_combined_PSD_CP3, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'CPz'))
                canal='CPz';
                if ismember(canal, good_channels)
                    left_combined_PSD_CPz= [left_combined_PSD_CPz, psd_left{c}];
                    right_combined_PSD_CPz= [right_combined_PSD_CPz, psd_right{c}];
                    left_combined_PSD_CPz= [left_combined_PSD_CPz, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_CPz= [right_combined_PSD_CPz, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'CP4'))
                canal='CP4';
                if ismember(canal, good_channels)
                    left_combined_PSD_CP4= [left_combined_PSD_CP4, psd_left{c}];
                    right_combined_PSD_CP4= [right_combined_PSD_CP4, psd_right{c}];
                    left_combined_PSD_CP4= [left_combined_PSD_CP4, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_CP4= [right_combined_PSD_CP4, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'FC4'))
                canal='FC4';
                if ismember(canal, good_channels)
                    left_combined_PSD_FC4= [left_combined_PSD_FC4, psd_left{c}];
                    right_combined_PSD_FC4= [right_combined_PSD_FC4, psd_right{c}];
                    left_combined_PSD_FC4= [left_combined_PSD_FC4, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_FC4= [right_combined_PSD_FC4, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'FCz'))
                canal='FCz';
                if ismember(canal, good_channels)
                    left_combined_PSD_FCz= [left_combined_PSD_FCz, psd_left{c}];
                    right_combined_PSD_FCz= [right_combined_PSD_FCz, psd_right{c}];
                    left_combined_PSD_FCz= [left_combined_PSD_FCz, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_FCz= [right_combined_PSD_FCz, NaN(length(psd_right{c}),1)];
                end
            elseif(strcmp((EEG.chanlocs(c).labels),'FC3'))
                canal='FC3';
                if ismember(canal, good_channels)
                    left_combined_PSD_FC3= [left_combined_PSD_FC3, psd_left{c}];
                    right_combined_PSD_FC3= [right_combined_PSD_FC3, psd_right{c}];
                    left_combined_PSD_FC3= [left_combined_PSD_FC3, NaN(length(psd_left{c}),1)];
                    right_combined_PSD_FC3= [right_combined_PSD_FC3, NaN(length(psd_right{c}),1)];
                end
            else
                continue
            end
        
        end
    
    end
    
    % Names of files to store ERD-ERS ratio in PSD
    leftexcelFileName =fullfile(folderPath,'ERD_ERS_left.xlsx') ;
    rightexcelFileName =fullfile(folderPath,'ERD_ERS_right.xlsx') ;
    
    for c = 1:length(good_channels)
    
        if(strcmp((EEG.chanlocs(c).labels),'C3'))
            canal='C3';
            if ismember(canal, good_channels)
                left_combinedPSD = left_combined_PSD_C3;
                right_combinedPSD = right_combined_PSD_C3;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'C4'))
            canal='C4';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_C4;
            right_combinedPSD = right_combined_PSD_C4;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'Cz'))
            canal='Cz';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_Cz;
            right_combinedPSD = right_combined_PSD_Cz;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'CP3'))
            canal='CP3';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_CP3;
            right_combinedPSD = right_combined_PSD_CP3;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'CPz'))
            canal='CPz';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_CPz;
            right_combinedPSD = right_combined_PSD_CPz;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'CP4'))
            canal='CP4';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_CP4;
            right_combinedPSD = right_combined_PSD_CP4;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'FC4'))
            canal='FC4';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_FC4;
            right_combinedPSD = right_combined_PSD_FC4;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'FCz'))
            canal='FCz';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_FCz;
            right_combinedPSD = right_combined_PSD_FCz;
            end
        elseif(strcmp((EEG.chanlocs(c).labels),'FC3'))
            canal='FC3';
            if ismember(canal, good_channels)
            left_combinedPSD = left_combined_PSD_FC3;
            right_combinedPSD = right_combined_PSD_FC3;
            end
        else
            continue
        end
        if ismember(canal, good_channels)
            % Create a sheet name for the channel
            sheetName = ['Channel_' canal];
            
            % Write the combined PSD matrix to the Excel file on the respective sheet
            writematrix(freq, leftexcelFileName, 'Sheet', sheetName, 'Range', 'A2');
            writecell(filename, leftexcelFileName, 'Sheet', sheetName, 'Range', 'B1');    
            writematrix(left_combinedPSD, leftexcelFileName, 'Sheet', sheetName, 'Range', 'B2');
            writematrix(freq, rightexcelFileName, 'Sheet', sheetName, 'Range', 'A2');
            writecell(filename, rightexcelFileName, 'Sheet', sheetName, 'Range', 'B1');
            writematrix(right_combinedPSD, rightexcelFileName, 'Sheet', sheetName, 'Range', 'B2');
        end
        close all
    
    end
end


%% Functions


function [psd_total,freq] = plot_spectrogram(eeg_data, fs, freq_range, title_str, channel_names)
    % Function to plot spectrogram
    [ ~,n_channels] = size(channel_names);
    n_rows = 3;
    n_cols = 3;
    
    figure;
    
    for i = 1:n_channels
        subplot(n_rows, n_cols, i);
        % Compute and plot the spectrogram with increased time resolution
        % Increasing the window length (e.g., to 512) and overlap (e.g., to 256)
        [~, freq, time, psd]=spectrogram(cell2mat(eeg_data(i)), [], [], [], fs, 'yaxis');
        % Plot the ERD/ERS ratio as a function of time and frequency
        surf(time, freq, 10*log10(abs(psd)), 'EdgeColor', 'none');
        power_limits = [-20, 20]; % Adjust these values as needed
        psd_total{i}=psd;
        view(2);
        axis tight;
        ylim(freq_range);
        xlabel('Time (s)');
        ylabel('Frequency (Hz)');
        zlabel('ERD/ERS Ratio (%)');
        title(sprintf('Channel %s', cell2mat(channel_names(i))));
        colorbar;
        clim(power_limits); % Set the color axis limits
    end
    sgtitle(title_str);
end

function smoothed_data = moving_average(data, window_size)
    % Function to perform a simple moving average smoothing
    %
    % Inputs:
    % - data: The data to be smoothed (vector)
    % - window_size: The size of the moving window (scalar, in samples)
    %
    % Outputs:
    % - smoothed_data: The smoothed data (vector)

    % Initialize the smoothed data
    smoothed_data = zeros(size(data));

    % Perform the moving average
    for i = 1:length(data)
        if i <= window_size
            % For the initial points, take the average from the start to the current point
            smoothed_data(i) = mean(data(1:i));
        else
            % For other points, take the average over the window_size
            smoothed_data(i) = mean(data(i-window_size+1:i));
        end
    end
end


function subject_data = select_good_channels(path_to_file)
    
    % Load the Excel file into a table
    data = readtable(path_to_file);
    
    % Extract the subject IDs (first column) and file names (second column)
    subject_ids = data{:, 1};
    file_names = data{:, 2};
    
    % Extract the channel data (assuming the good/bad markers are in columns 3 to 20)
    channel_data = data(:, 3:20);
    
    % Initialize a structure to store files and good channels by subject
    subject_data = struct();
    
    % Get the names of the channels from the column headers
    channel_names = data.Properties.VariableNames(3:20);
    
    % Loop through each subject and group files and good channels by subject
    for i = 1:length(subject_ids)
        % Split the subject ID by '_'
        split_subject = split(subject_ids{i}, '_');
        
        % Extract the subject identifier (e.g., 'S3')
        subject = split_subject{1};
        
        % Find the good channels (marked as 1 or 2)
        good_channels = channel_names(ismember(table2array(channel_data(i, :)), [1, 2]));
    
        modified_channel_names = cell(1,length(good_channels));
    
        % Loop through each channel name and extract the part before '_'
        for j = 1:length(good_channels)
            % Split the channel name by '_'
            split_name = split(good_channels{j}, '_');
            
            % Store the part before the '_'
            modified_channel_names{j} = char(split_name(1));  % Convert string to char
        end
    
        % Eliminate duplicate channel names
        unique_channel_names = unique(modified_channel_names);
    
        
        % Create a struct for this file entry
        file_entry = struct();
        file_entry.file = file_names{i};  % Store the file name
        file_entry.good_channels = unique_channel_names;  % Store the good channels
        
        % Check if the subject already exists in the structure
        if isfield(subject_data, subject)
            % Append the new file entry to the existing list for this subject
            subject_data.(subject)(end+1) = file_entry;
        else
            % Initialize a new list for this subject with the current file entry
            subject_data.(subject) = file_entry;
        end
    end
end


