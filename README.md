<br />
<p align="center">
  <a href="https://github.com/CognitiveNeuroLab/matlab-plot-to-powerpoint/">
    <img src="images/logo.jpeg" alt="Logo" width="160" height="160">
  </a> 

<h3 align="center">Plotting matlab figures in powerpoint</h3>

<h4 align="center"> This repo has 1 script that allows you to plot EEG data using Matlab, but instead save it in a Powerpoint. Making it easy for presentations where doing this manually takes hours </h4>

**Table of Contents**
  
1. [About the project](#about-the-project)
    - [Built With](#built-with) 
    - [Parts you need to customize](#parts-you-need-to-customize) 
2. [License](#license)  
3. [Contact](#contact)  


## About The Project

Creating plots and putting these into a powerpoint to present an ongoing project and go over data always takes long. So instead of that this script does it for you. The code used is native to EEGLAB but can easly be changed to any load/plotting function for a different toolbox. If you already have saved the plotted files, you can just uncommand the plotting part and only use the PowerPoint creation part.

### Built With

* [Matlab 2019b)

### Parts you need to customize  
  
Before starting you need to change a couple of things:  
   
``` matlab
home_path = '\\data.einsteinmed.org\users\MoBI\MOBI ASD\SIFT\data\'; %if you need to load data
figure_path='\\data.einsteinmed.org\users\MoBI\MOBI ASD\SIFT\figure\'; %where to load/save figures
save_path='\\data.einsteinmed.org\users\MoBI\MOBI ASD\SIFT\'; %where to save the excel figures
subject_list  = {'10049' '10056' '10463'};% This will be the amount of particpant's data you want to plot  
PowerPoint_title = 'All ICs MoBI ASD.pptx';
```
   
The `subject_list` variable in the case of the base script has ID numbers because it's going to run through different peoples data. But this could also be channels names, if that is what you want to plot. 
  
In line 17 you can change the titles of the first slide
In line 18 you can change the subtitle of the first slide
In line 25 and 27 you can do the same for the second slide

In line 34-40 you load data (EEGLAB style) and plot it. If you use a different plot function, or load data using something like fieldtrip, you should change or uncommand this. 

In line 41-46 we create a base slide for each participant. You can either modify this or uncommand it all together. 

### Example powerpoint  
  
See the powerpoint called "All ICs MoBI ASD" to see how the result of this plotting works
  
       
## License

Distributed under the MIT License. See [LICENSE](https://github.com/CognitiveNeuroLab/matlab-plot-to-powerpoint/blob/master/LICENSE.txt) for more information.

## Contact

Douwe Horsthuis - [@douwejhorsthuis](https://twitter.com/douwejhorsthuis) - douwehorsthuis@gmail.com - www.douwehorsthuis.com

Project Link: [https://github.com/CognitiveNeuroLab/matlab-plot-to-powerpoint/](https://github.com/CognitiveNeuroLab/matlab-plot-to-powerpoint/)
