# postal_data_analysis.py

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Set style for better visualization
sns.set(style="whitegrid")

def load_and_clean_data(file_path):
    """Load and clean the dataset"""
    try:
        # Load the dataset
        df = pd.read_excel(file_path)
        
        # Data cleaning
        df_cleaned = df.copy()
        
        # Drop missing values
        df_cleaned = df_cleaned.dropna()
        
        # Drop duplicates
        df_cleaned = df_cleaned.drop_duplicates()
        
        # Standardize column names and string values
        df_cleaned.columns = df_cleaned.columns.str.strip().str.lower()
        df_cleaned = df_cleaned.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        
        # Verify required columns exist
        required_columns = ['statename', 'district', 'officetype', 'regionname', 'delivery']
        for col in required_columns:
            if col not in df_cleaned.columns:
                raise ValueError(f"Required column '{col}' not found in dataset")
                
        return df_cleaned
        
    except Exception as e:
        print(f"Error loading or cleaning data: {str(e)}")
        return None

def generate_summary_stats(df):
    """Generate summary statistics"""
    print("\n=== BASIC SUMMARY ===")
    print("Total Entries After Cleaning:", len(df))
    print("Unique States:", df['statename'].nunique())
    print("Unique Districts:", df['district'].nunique())
    print("Unique Office Types:", df['officetype'].unique())

def create_visualizations(df):
    """Create and save visualizations"""
    
    # Visualization 1: Post Offices by Region
    plt.figure(figsize=(12, 8))  # Increased figure size
    df['regionname'].value_counts().plot(kind='barh', color='skyblue')
    plt.title("Number of Post Offices by Region", pad=20)
    plt.xlabel("Number of Post Offices")
    plt.ylabel("Region")
    plt.tight_layout(pad=2.0)  # Added padding
    plt.savefig("region_wise_postoffices.png", bbox_inches='tight')
    plt.close()
    
    # Visualization 2: Post Office Types
    plt.figure(figsize=(10, 6))  # Increased figure size
    df['officetype'].value_counts().plot(kind='bar', color='orange')
    plt.title("Distribution of Post Office Types", pad=20)
    plt.xlabel("Office Type")
    plt.ylabel("Count")
    plt.tight_layout(pad=2.0)
    plt.savefig("postoffice_types.png", bbox_inches='tight')
    plt.close()
    
    # Visualization 3: Delivery vs Non-Delivery per State
    plt.figure(figsize=(16, 10))  # Increased figure size
    delivery_counts = df.groupby(['statename', 'delivery']).size().unstack().fillna(0)
    ax = delivery_counts.plot(kind='bar', stacked=True, colormap='Set2')
    plt.title("Delivery vs Non-Delivery Post Offices by State", pad=20)
    plt.ylabel("Number of Post Offices")
    plt.xlabel("State")
    plt.xticks(rotation=45, ha='right')  # Better rotation and alignment
    plt.tight_layout(pad=3.0)  # Increased padding
    plt.savefig("delivery_vs_non_delivery.png", bbox_inches='tight')
    plt.close()
    
    # Visualization 4: Top 10 States with Most Post Offices
    plt.figure(figsize=(12, 6))
    state_counts = df['statename'].value_counts().head(10)
    ax = state_counts.plot(kind='bar', color='green')
    plt.title("Top 10 States with Most Post Offices", pad=20)
    plt.xlabel("State")
    plt.ylabel("Number of Post Offices")
    plt.xticks(rotation=45, ha='right')
    
    # Add value labels on top of bars
    for p in ax.patches:
        ax.annotate(f"{int(p.get_height())}", 
                   (p.get_x() + p.get_width() / 2., p.get_height()), 
                   ha='center', va='center', 
                   xytext=(0, 5), 
                   textcoords='offset points')
    
    plt.tight_layout(pad=2.0)
    plt.savefig("top10_states_postoffices.png", bbox_inches='tight')
    plt.close()
    
    # Visualization 5: Top 10 Districts (Heatmap)
    plt.figure(figsize=(12, 3))  # Adjusted figure size
    top_districts = df['district'].value_counts().head(10)
    sns.heatmap(top_districts.to_frame().T, annot=True, cmap="YlGnBu", fmt=".0f", cbar=False)
    plt.title("Top 10 Districts with Most Post Offices", pad=20)
    plt.tight_layout(pad=1.0)
    plt.savefig("top10_districts_heatmap.png", bbox_inches='tight')
    plt.close()

def main():
    # File path - update this with your actual path
    file_path = r"C:\Users\Asus\Downloads\datasetGPTkeLiye.xlsx"
    
    # Load and clean data
    print("Loading and cleaning data...")
    df_cleaned = load_and_clean_data(file_path)
    
    if df_cleaned is not None:
        # Generate summary statistics
        generate_summary_stats(df_cleaned)
        
        # Create visualizations
        print("\nCreating visualizations...")
        create_visualizations(df_cleaned)
        
        # Save cleaned dataset
        df_cleaned.to_excel("cleaned_postal_dataset.xlsx", index=False)
        print("\nCleaned dataset saved as 'cleaned_postal_dataset.xlsx'")
        
        print("\nAll visualizations saved successfully!")
    else:
        print("Failed to process data. Please check the input file and try again.")

if __name__ == "__main__":
    main()