#  Plagiarism Validation Tool  

 **A C#-based tool for optimizing software review processes by filtering insignificant matches and eliminating cyclical redundancy.**  
By prioritizing the most relevant similarities, this tool reduces evaluator workload and maximizes efficiency.  

---

##  Project Description  
This project applies **algorithmic techniques** to a dataset loaded from an Excel file, with the primary goal of detecting plagiarism through:  
 Identifying **connected components** within the similarity graph.  
 Computing **average similarity weights** for better analysis.  
 Constructing **maximum spanning trees (MST)** to highlight significant relationships.  
 Exporting processed results into **Excel files** for further evaluation.  

---

##  Project Functionality  

### 1 **Loading Data (LOAD_DATA)**  
 Reads input from an **Excel file** containing file paths and similarity measures.  
 Constructs a **graph representation** of file similarities for analysis.  

### 2 **Creating Connected Components (CreateConnectedComponents)**  
 Uses **Depth-First Search (DFS)** to identify clusters of related documents.  
 Returns a list of components, each containing closely related files.  

### 3 **Calculating Average Weight (CalculateAverageWeight)**  
 Computes the **average similarity weight** within each connected component.  
 Helps in assessing overall similarity trends in different document clusters.  

### 4 **Graph Segmentation (GetEdgesInComponent)**  
 Separates **graph edges** based on identified connected components.  
 Organizes similarity data for **efficient MST construction**.  

### 5 **Finding Maximum Spanning Tree (FindMaxSpanningTree)**  
 Constructs the **Maximum Spanning Tree (MST)** for each component.  
 Uses a **modified Kruskal’s algorithm** to extract the most relevant similarity paths.  

### 6 **Exporting Results to Excel (ConvertDictionaryToExcel, ConvertListToExcel)**  
 Converts processed results into **Excel format** for further visualization and evaluation.  
 Supports **detailed metadata export** for component statistics and MST analysis.  

---

##  **Performance Metrics (Hard Test Cases)**  

| Functionality | Best Performance Time |
|--------------|-----------------------|
|  **Read Excel & Construct Graph** | 972 ms |
|  **Create Connected Components** | 11 ms |
|  **Calculate Average Weight** | 127 ms |
|  **Prepare Graph & Find MST** | 250 ms |
|  **Convert Group Statistics to Excel** | 157 ms |
|  **Convert Matching Pairs to Excel** | 1007 ms |
|  **State File Generation & Save** | 295 ms |
|  **MST Generation & Save (Kruskal’s Algorithm)** | 1757 ms |
|  **Total Best Performance Time** | **3750 ms** |

---

##  Technologies Used  
- **Programming Language:** C#  
- **Algorithms:** Graph Theory, DFS, Kruskal’s Algorithm  
- **Data Handling:** Excel Import/Export  
- **Framework:** .NET Framework / .NET Core  


