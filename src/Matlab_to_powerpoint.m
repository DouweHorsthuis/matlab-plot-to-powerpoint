%Creating a powerpoint with all ICs
clear variables
%% things that you need to fill out before you can run it
home_path = '\\data.einsteinmed.org\users\MoBI\MOBI ASD\SIFT\data\'; %if you need to load data
figure_path='\\data.einsteinmed.org\users\MoBI\MOBI ASD\SIFT\figure\'; %where to load/save figures
save_path='\\data.einsteinmed.org\users\MoBI\MOBI ASD\SIFT\'; %where to save the excel figures
%
  subject_list  = {'10049' '10056' '10463'};
 PowerPoint_title = 'All ICs MoBI ASD.pptx';
%% Power point start
disp('Creating Powerpoint');
import mlreportgen.ppt.*
ppt = Presentation(PowerPoint_title);
% title slide
disp('Creating Title slide');
titleSlide = add(ppt,'Title Slide');
replace(titleSlide,'Title','All ICs MoBI ASD');
subtitle=Text('Douwe 5/13/2022');
subtitle.FontSize= '16';
subtitleText = Paragraph(subtitle);
replace(titleSlide,'Subtitle',subtitleText);
% Summary slide
disp('Creating summary page');
textSlide = add(ppt,'Title and Content'); %adding slide (Title and Content is a text based template)
titleText = Paragraph('Summary'); % adding title
replace(textSlide,'Title',titleText);
replace(textSlide,'Content',{'Showing all ICs per participant for Hits',... %adding the content
    'Plots contain:',{'Time frequency for that IC', 'Headmap', 'ITC',}}); %{between these bracets youindent  }
for s=1:length(subject_list)
    fprintf('\n******\nProcessing subject %s\n******\n\n', subject_list{s});
    data_path  = [home_path subject_list{s} '\\'];
    % Load original dataset
    fprintf('\n\n\n**** %s: Loading dataset ****\n\n\n', subject_list{s});
    EEG = pop_loadset('filename', [subject_list{s} '_ready.set'], 'filepath', data_path);
    for ic_n=1:size(EEG.icaweights,1)
        % choose here whatever you want to plot, if you already have the figures saved somewhere, uncommand
        figure; pop_newtimef( EEG, 0, ic_n, [EEG.times(1)  EEG.times(end)], [3         0.8] , 'topovec', EEG.icawinv(:,ic_n), 'elocs', EEG.chanlocs, 'chaninfo', EEG.chaninfo, 'caption', ['IC ' num2str(ic_n)], 'baseline',[0], 'plotphase', 'off', 'padratio', 1, 'winsize', 64);
        print([figure_path subject_list{s} num2str(ic_n)], '-dpng' ,'-r300');
        close all
    end
    %% Participant 1st slide
    pictureSlide = add(ppt,'Title and Picture'); % adding slide (picture template)
    titleText = Paragraph(subject_list{s}); % adding title
    plot = Picture([figure_path subject_list{s} '_ICs_dipfitted.png']);
    replace(pictureSlide,'Title',titleText); %adding final title
    replace(pictureSlide,'Picture',plot);  % adding final plot
    %% adding individual 
    for ic_n=1:size(EEG.icaweights,1) %%size(EEG.icaweights,1) = the amount of ICs of that participant
        disp(['**** Working on IC ' , num2str(ic_n) '/' num2str(size(EEG.icaweights,1)) '  ****']);
        pictureSlide = add(ppt,'Title and Picture'); % adding slide (picture template)
        titleText = Paragraph(subject_list{s}); % adding title
        plot = Picture([figure_path subject_list{s} num2str(ic_n) '.png']);
        replace(pictureSlide,'Title',titleText); %adding final title
        replace(pictureSlide,'Picture',plot);  % adding final plot
    end
    %% deleting the figures, so that you don't end up with a full folder
    for ic_n=1:size(EEG.icaweights,1)
        delete([figure_path subject_list{s} num2str(ic_n)], '-dpng' ,'-r300');
    end
end
%% Saving PPT
close(ppt); % saves it

