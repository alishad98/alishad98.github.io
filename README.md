# alishad98-portfolio

#Data Analyst and Project Manager

#### Technical Skills: PowerBI, Python, SQL, Snowflake, ERP Systems (SAP)

## Education
- M.Sc, Sutainable Energy | Imperial College London (_September 2021_)								       		
- BEng, Mechanical Engineering| Cardiff University (_July 2019_)	 			        		

## Work Experience
**Cost Analyst - Budget Manager II @ Mercedes-AMG Petronas Formula One Team(_December 2023 - Present_)**
- Uncovered and corrected missing step in production data pipeline which impacted over 70% of active accounts
- Redeveloped loan originations model which resulted in 50% improvement in model performance and saving 1 million dollars in potential losses

**Cost Analyst - Inventory Budgets I @ Mercedes-AMG Petronas Formula One Team (_November 2022 - December 2023_**

**Data Analyst @ Emitwise (_October 2021 - October 2022_)**
- Conducted data collection, processing, and analysis for novel study evaluating the impact of over 300 biometrics variables on human performance in hyper-realistic, live-fire training scenarios
- Applied unsupervised deep learning approaches to longitudinal ICU data to discover novel sepsis sub-phenotypes

## Projects
### Formula One Budget and Inventoy Insights PowerBI Dashboard

Dsigned dashboard to streamline budget tracking and provide real-time insights into cost variances for the Formula One programme, integrating data from multiple corporate platforms, including **Snowflake**, **Python** and **SAP**. This dashboard is actively used by and presented to stakeholders, including C-suite executives, to make data-driven decisions, monitor cost performance, and optimise inventory processes with the objective to indirectly impact car performance and provide more available spend to develop the car.

#### Key Features and Functionality
- **Consolidated Data Sources:** Links SAP and other corporate platforms into a single view, centralising all budget and inventory data for seamless analysis.
- **Budget Tracking:** Displays current spend against allocated budgets for all car inventory, categorised by group, enabling stakeholders to monitor financial health at a glance.
- **Detailed Inventory Insights:** Tracks inventory value ordered per budget group, Highlights delivered inventory vs. remaining value, and provides a split between internal manufacturing and external purchasing, offering clarity on operations.
- **Visualisation of Future Orders:** Right-hand graphs display order value by delivery date, highlighting the status of future orders, Identifies variances between planned and actual spending.
- **Forecasting Capabilities:** Includes a dynamic running forecast, providing stakeholders with predictive insights for proactive decision-making.

#### Development Process
This dashboard was developed based on stakeholder requirements and through in-depth analysis of corporate data. It involved identifying and connecting datasets across multiple dashboards.
Building relationships between data sources to ensure accuracy and consistency.
Designing user-centric visuals and reports tailored to meet the needs of cross-functional teams.





### Python Script for SAP Data Automation

Developed a Python-based automation script to connect to SAP and streamline the extraction, processing, and integration of live production data. Designed to run twice daily, the script ensures stakeholders have access to up-to-date information critical for tracking production orders, financial settlements, and part costs.

#### Key Features and Functionality
- **Automated SAP Interaction:** Utilised SAP recording and playback coding to log in, access key transactions, and export essential data.
Extracted data on production orders, time bookings, part costs, and the financial settlement of production orders to budget codes.
- **Data Processing and Management:** Leveraged Python libraries, including pandas, openpyxl, and pyperclip, for efficient data manipulation. Ensured only new data was incorporated into recent exports, avoiding duplication and maintaining data integrity.
- **Excel File Manipulation:** Processed exported SAP data to prepare refined, actionable outputs for downstream users.
Development Process
- **Tools and Libraries:** I worked extensively with Python libraries such as pandas for data manipulation, openpyxl for Excel integration, and pyperclip for clipboard interactions to ensure a seamless automation process.
- **Customisation:** Script functionality was tailored based on stakeholder requirements to optimise daily operations and align with financial and production tracking needs.

#### Impact
Reduced manual effort by automating twice-daily SAP interactions, freeing up significant time for team members.
Provided real-time insights into production and financial data, enabling faster decision-making and improved operational efficiency.
Enhanced data accuracy and consistency by ensuring clean and non-duplicated records.



```ruby
require 'redcarpet'
markdown = Redcarpet.new("Hello World!")
puts markdown.to_html
```

### Process Optimisation Projects

#### Power Automate
To streamline workflows and reduce manual intervention, I developed a Power Automate flow chain that optimised email-based communication and data management processes across teams. This automation achieved the following:
- **Reduced Manual Effort:** Automated the distribution of critical reports and reminders, ensuring timely updates without requiring manual follow-ups.
- **Improved Team Efficiency:** Enabled seamless coordination across departments by automatically routing emails to relevant stakeholders based on predefined conditions.
- **Error Mitigation:** Eliminated manual errors in repetitive tasks by ensuring consistent data handling and distribution.

Key Approaches included:
- **Conditional Logic:** Incorporated rules to send tailored emails based on triggers like deadlines or changes in data.
- **Integration with Other Tools:** Linked with SharePoint for file storage and Outlook for email delivery, ensuring all information was accessible and actionable.
- **Real-Time Updates:** Included dynamic updates to accommodate changes in data, ensuring stakeholders always had the most current information.

Screenshots of anonymised workflows and process diagrams are included to provide insight into the design and functionality of this project.

#### Business Process Improvement

As part of my ongoing efforts to optimise business processes, I developed an Authorised Document Flowchart to streamline and clarify the document approval process across the company. The objective was to reduce bottlenecks, increase efficiency, and ensure compliance with internal policies.

Key aspects of the flowchart include:
- **Clear Workflow Visualisation:** Designed a step-by-step process that identifies each stage in the document approval journey, from initiation to final authorisation. This ensured that team members understood where their responsibilities lay at every stage.
- **Role-Based Access:** Incorporated role-specific permissions, ensuring that only authorised individuals could approve or modify documents at various stages, reducing the risk of errors or compliance issues.
- **Automated Trigger Points:** Integrated automatic notifications and reminders for team members to ensure timely progress through the approval pipeline.
- **Efficient Escalation:** Defined clear escalation paths to prevent delays in approval, ensuring that critical documents were fast-tracked when necessary.

The document and process diagram served as a visual reference and also as the foundation for automating the document management process using Power Automate. This combined approach helped significantly reduce delays, and improve transparency and cross-functional collaboration. An anonymised version of the flowchart is included to provide an overview of the logic and design.




### Financial Analysis Reporting 
As part of my role in financial reporting, I developed and regularly sent out comprehensive Inventory & Cost Management Reports to Director-level and C-suite executives across the business, focusing on key financial insights related to inventory tracking and cost variances. The reports served to inform senior management and key stakeholders on current and projected financial performance.

Key elements of this project include:
- **Inventory Data Extraction & Analysis:** Utilised advanced Excel techniques and SQL to extract, clean, and analyse large volumes of data from various sources, including SAP, to track the inventory value, ordered quantities, and delivery statuses.
- **Cost Variance Analysis:** Created detailed reports that identified variances between budgeted and actual costs for car inventory and Formula 1 program components, highlighting discrepancies and opportunities for cost-saving.
- **Fuel Line Graph:** Developed an in-depth fuel line graph to visualise the ongoing expenditure and project future costs. This graph was integral to senior leadership's understanding of financial trends, helping to optimise spending and forecast future budget needs.
- **Dynamic Reporting:** Designed a report that could be updated easily with new data, ensuring stakeholders always had the most current information. The report was structured to allow quick decision-making and ongoing strategic planning.

The reports and associated visuals (such as the fuel line graph) were key to driving financial decisions and were frequently used by senior leadership for cost optimisation and budgeting purposes. An anonymised version of the report and the associated fuel line graph is included to showcase the type of data and insights provided.



