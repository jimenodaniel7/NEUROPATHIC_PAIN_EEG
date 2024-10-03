% EEGLAB history file generated on the 16-Jun-2023
% ------------------------------------------------
function [EEG, ERP, ERP_TFFT, ERP_EFFT]=erp_bin(name_set, filePath, pwd)


%eeglab
EEG = pop_loadset('filename',name_set,'filepath',filePath);
EEG = pop_selectevent( EEG, 'type',{'IMGTASK_LA','IMGTASK_RA','boundary'},'deleteevents','on');
EEG  = pop_editeventlist( EEG , 'AlphanumericCleaning', 'on', 'BoundaryNumeric', { -99}, 'BoundaryString', { 'boundary' }, 'List', strcat(pwd,'/jk.txt'), 'SendEL2', 'EEG', 'UpdateEEG', 'codelabel', 'Warning', 'on' ); % GUI: 16-Jun-2023 00:11:31 

EEG = pop_epochbin( EEG , [-1000.0  2000.0],  'pre'); % GUI: 16-Jun-2023 00:12:04
%EEG = pop_editset(EEG, 'setname', 'S1_Chan');


EEG.data = EEG.data.^2;

ERP = pop_averager( EEG , 'Criterion', 'good', 'DQ_custom_wins', 0, 'DQ_flag', 1, 'DQ_preavg_txt', 0, 'ExcludeBoundary', 'on', 'SEM',...
 'on' );
ERP_TFFT = pop_averager( EEG , 'Compute', 'TFFT', 'Criterion', 'good', 'DQ_custom_wins', 0, 'DQ_flag', 1, 'DQ_preavg_txt', 0, 'ExcludeBoundary',...
  'on', 'SEM', 'on' );
 ERP_EFFT = pop_averager( EEG , 'Compute', 'EFFT', 'Criterion', 'good', 'DQ_custom_wins', 0, 'DQ_flag', 1, 'DQ_preavg_txt',...
  0, 'ExcludeBoundary', 'on', 'SEM', 'on' );

ERP = pop_ploterps( ERP, [ 1 2],  1:length(EEG.chanlocs) , 'AutoYlim', 'on', 'Axsize', [ 0.05 0.08], 'BinNum', 'on', 'Blc', 'pre', 'Box', [ 3 3], 'ChLabel',...
 'on', 'FontSizeChan',  10, 'FontSizeLeg',  12, 'FontSizeTicks',  10, 'LegPos', 'bottom', 'Linespec', {'k-' , 'r-' }, 'LineWidth',  1, 'Maximize',...
 'on', 'Position', [ 68.6429 15.0714 106.857 31.9286], 'Style', 'Classic', 'Tag', 'ERP_figure', 'Transparency',  0, 'xscale',...
 [ -1000.0 1996.0   -800:400:1600 ], 'YDir', 'normal' );
%ERP = pop_savemyerp(ERP, 'erpname', 'S1_ERPs', 'filename', 'S1_ERPs.erp', 'filepath', '/Users/vanesasotoleon/Downloads/ERD_ERS/ERD_ERS_ERP', 'warning', 'on');  


% ALLERP = ERP;
% ALLEEG=EEG;
% [ALLEEG, EEG, CURRENTSET] = eeg_store( ALLEEG, EEG, 0 );
% 
% erplab redraw
% eeglab redraw;

