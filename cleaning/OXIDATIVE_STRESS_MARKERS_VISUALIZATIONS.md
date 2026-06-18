# Data Visualization Report of Oxidative Stress Markers

This report summarizes and visualizes the results of the biochemical assays for Clove and Bay Leaf treatments. It compiles the bar charts, box plots, correlation matrix, and radar profile charts to describe group trends and statistical variances.

## Banned Charts Explanation

In biochemical pharmacology, certain chart types are omitted because they represent data distributions inaccurately.
* **Pie Charts**: Pie charts illustrate proportions of a whole summing to 100%. Because enzyme activities (SOD, CAT, GPx) and marker concentrations represent independent continuous parameters, dividing them into pie segments is scientifically incorrect.
* **Histograms**: Histograms illustrate data frequency distributions. Due to small cohort sizes (n = 3 or n = 5) in standard rodent models, histograms do not provide meaningful distributions. Standard box-and-whisker plots overlaid with individual data points are used instead to show data dispersion.

---

## 1. Summary Bar Charts (Mean ± SEM)

Grouped bar charts illustrate the mean value of each biochemical parameter across all groups. Standard error of the mean (SEM) is represented as error bars.

### Clove Assay Markers
The bar chart below summarizes the seven biochemical parameters measured during the Clove assay.

![Clove Assay Summary Bar Chart](graphs/clove_markers_bar_chart.png)

### Bay Leaf Assay Markers
The bar chart below summarizes the seven parameters measured during the Bay Leaf assay.

![Bay Leaf Assay Summary Bar Chart](graphs/bay_leaf_markers_bar_chart.png)

---

## 2. Data Distribution Box Plots

Box-and-whisker plots display individual data points to show the exact distribution, median, and range within each group. This format is preferred to examine biological outliers.

### Superoxide Dismutase (SOD) Distributions

#### Clove SOD Distribution
![Clove SOD Box Plot](graphs/clove_sod_boxplot.png)

#### Bay Leaf SOD Distribution
![Bay Leaf SOD Box Plot](graphs/bay_leaf_sod_boxplot.png)

### Malondialdehyde (MDA) Distributions

#### Clove MDA Distribution
![Clove MDA Box Plot](graphs/clove_mda_boxplot.png)

#### Bay Leaf MDA Distribution
![Bay Leaf MDA Box Plot](graphs/bay_leaf_mda_boxplot.png)

---

## 3. Parameter Correlations

A Pearson correlation matrix calculates the relationship between each biochemical marker across all combined experimental groups.

![Pearson Correlation Heatmap](graphs/markers_correlation_heatmap.png)

### Correlation Observations
The correlation heatmap displays a strong negative correlation between antioxidant enzyme levels and lipid peroxidation (MDA). For example, increased SOD and CAT activities correlate with decreased MDA levels, confirming the protective mechanism against cell membrane damage.

---

## 4. Multi-Dimensional Profiling (Radar Chart)

The radar chart maps the normalized means of all seven parameters on a single radial axis. This provides a profile of how closely a treatment group matches either the healthy control or the model control.

![Clove Assay Radar Chart](graphs/clove_markers_radar_chart.png)

### Radar Observations
The radar chart contrasts the protective profile of the treatments. It shows that the Normal Control group maintains high antioxidant levels, whereas the Model Control group displays significant depletion. The overlap between Clove and Model Control groups illustrates the lack of recovery in the Clove treated cohort.
