# NEUROPATHIC_PAIN_EEG
This is a repository with MATLAB scripts that extract and analyze information from EEG files from a Hero Bitbrain EEG device. This data was collected in a series of experiments develos in the FENNSI investigation group in the National Hospital of Paraplegics in Toledo, Spain. In the experiments, a subject had to a series of tasks: 2 minutes with his eyes open trying not to think anything, 2 minutes doing the same but with the eyes close, 10 minutes of motor imaginary task (thinking about moving your hand without actually doing it) and lastly another 2 minutes wit the eyes open. During the imaginary task, the patient had to imaging a movement with the left or the right hand, depending on the direction of the arrow that showed the software used in the experiment.
This experiment was made with people that suffered neuropathic pain from a spinal cord injury. The final goal was to find a relation between the pain and the cortex activity captured with the EEG device.

The files have to be ran in a specific order:

1. bandpass_asr.m
2. all_excels_afterasr_withoutIT.m
3. trials.m
4. calc_psd.m

The first one, bandpass_asr.m, loads the data from the files provided by the EEG device and filter the signals. First, a bandpass filter is used, and then the ASR method. It returns the data filtered in csv format.
The second file calculates the PSD from every channel in the data and extracts the power of each of the five frequency bands of activity recorded in the brain (alpha, beta, gamma, delta, theta).
The third file, trials.m, calculates the PSD from the trials in the motor imaginary task. Each trial was marked by the arrow that showed in the program.
The last file, calc_psd.m, extracts the power of the frequency bands of the PSD of the motor task. First, it calculates the mean of all the PSD of the trials of each hand, and then it extracts the power from that signals.

## New additions 3-10-2024

With the processed EEG files, the new script "test_ERPanalysis_ERDERS_spectrogram.m" does the following steps:
1. Compute ERP of the trials of each experiment
2. Compute ERD-ERS ratio of the ERP and show the spectrogram

The ERD-ERS ratio is computed with the following formula:

$$
ERDERS% = 100 \times \frac{smoothed\_ERP - \mu_{baseline}}{\mu_{baseline}}
$$

The MATLAB function "erp_bin.m" computes the ERP of the experiments.
