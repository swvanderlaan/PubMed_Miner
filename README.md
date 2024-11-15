# [PubMed Miner](https://github.com/swvanderlaan/PubMed_Miner)<img align="right" height="200" src=images/FullLogo_Transparent.png>

[![Languages](https://skillicons.dev/icons?i=py)](https://skillicons.dev) 

Mine PubMed for publication at the Central Diagnostics Laboratory of the Division Laboratories, Pharmacy and Biomedical genetics at the University Medical Center Utrecht, Utrecht University, Utrecht, The Netherlands.

To install follow these instructions. First, clone the repository to a directory of your choice:

```
cd to_directory_of_choice
git clone 

```

Next, create a new conda environment:

```
mamba create --name pubmedminer python=3.9
```

Activate the environment:

```
conda activate pubmedminer
```

Install the required packages:

```
mamba install biopython python-docx matplotlib numpy pandas
```

And some `pip` packages:

```
pip install xlsxwriter
```

Finally, run the script:

```
python pubmed_miner.py --email your_mail@whatever.com --verbose --year 2023-2024 --names "last_name IN"
```


<a href='https://www.umcutrecht.nl/en/centraal-diagnostisch-laboratorium'><img src='images/UMCU_2019_logo_liggend_rgb.png' align="center" height="75" /></a> 

#### Changes log
    
    _Version:_      v1.0.0</br>
    _Last update:_  2024-11-14</br>
    _Written by:_   Sander W. van der Laan (s.w.vanderlaan-2[at]umcutrecht.nl).
    
    **MoSCoW To-Do List**
    The things we Must, Should, Could, and Would have given the time we have.
    _M_

    _S_

    _C_

    _W_

    **Changes log**
    * v1.0.0 Initial version.

--------------

#### MIT License
##### Copyright (c) 1979-2024. Sander W. van der Laan | s.w.vanderlaan [at] gmail [dot] com | https://vanderlaanand.science.