import pandas as pd
from typing import Dict, Tuple, Optional, Any, List
from dataclasses import dataclass
from datetime import datetime



@dataclass
class ComparisonData:
    """Class to hold comparison analysis results"""
    recovered_hotels: pd.DataFrame
    new_missing_hotels: pd.DataFrame
    need_assistance_hotels: pd.DataFrame
    
    def to_dict(self) -> Dict[str, Dict[str, int]]:
        """Convert comparison data to dictionary format"""
        return {
            'recovered': {
                'count': len(self.recovered_hotels),
                'hotels': self.recovered_hotels.to_dict('records')
            },
            'new_missing': {
                'count': len(self.new_missing_hotels),
                'hotels': self.new_missing_hotels.to_dict('records')
            },
            'need_assistance': {
                'count': len(self.need_assistance_hotels),
                'hotels': self.need_assistance_hotels.to_dict('records')
            }
        }

@dataclass
class ProcessedData:
    """Class to hold all processed data"""
    missing: pd.DataFrame
    received: pd.DataFrame
    emails: pd.DataFrame
    comparison: Optional[ComparisonData] = None

@dataclass
class InsightResult:
    """Class to hold analysis results"""
    success: bool
    message: str
    data: Optional[Dict[str, Any]] = None

def load_excel_sheet(file: Any, sheet_name: str, column_names: List[str], skip_rows: int = 3) -> pd.DataFrame:
    """
    Load and process an Excel sheet with standardized processing
    
    Args:
        file: Excel file object
        sheet_name: Name of the sheet to load
        column_names: List of column names to use
        skip_rows: Number of rows to skip at start
    
    Returns:
        Processed DataFrame
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        df = df.iloc[skip_rows:].reset_index(drop=True)
        df = df.dropna(axis=1, how='all')
        
        actual_cols = min(len(df.columns), len(column_names))
        df.columns = column_names[:actual_cols]
        
        if 'business_date' in df.columns:
            df['business_date'] = pd.to_datetime(df['business_date'], format='%Y-%m-%d', errors='coerce')
        if 'days_missing' in df.columns:
            df['days_missing'] = pd.to_numeric(df['days_missing'], errors='coerce')
        
        return df.dropna(subset=['prop_code'])
    except Exception as e:
        raise Exception(f"Error loading Excel sheet {sheet_name}: {str(e)}")

def load_and_prepare_data(
    missing_file: Any,
    emails_file: Any,
    previous_audit_file: Optional[Any] = None
) -> Tuple[Optional[ProcessedData], Optional[str]]:
    """
    Load and prepare all required data files
    
    Args:
        missing_file: Current audit Excel file
        emails_file: Emails list Excel file
        previous_audit_file: Optional previous audit Excel file
    
    Returns:
        Tuple of (ProcessedData object, error message if any)
    """
    try:
        column_names = ['prop_code', 'hotel', 'file_name', 'business_date', 'dow', 'days_missing', 'status']
        
        # Load current audit data
        missing = load_excel_sheet(missing_file, 'Missing', column_names)
        received = load_excel_sheet(missing_file, 'Received', column_names)
        
        # Load and process emails
        emails = pd.read_excel(emails_file, header=0)
        emails.rename(columns={
            emails.columns[0]: 'Hotels',
            emails.columns[1]: 'Email'
        }, inplace=True)
        
        comparison_data = None
        if previous_audit_file is not None:
            comparison_data = generate_comparison_data(
                missing,
                load_excel_sheet(previous_audit_file, 'Missing', column_names)
            )
        
        return ProcessedData(missing, received, emails, comparison_data), None
    except Exception as e:
        return None, f"Error loading data: {str(e)}"

def generate_comparison_data(current_missing: pd.DataFrame, previous_missing: pd.DataFrame) -> ComparisonData:
    """
    Generate comparison analysis between current and previous audit data
    
    Args:
        current_missing: Current missing files DataFrame
        previous_missing: Previous missing files DataFrame
    
    Returns:
        ComparisonData object containing analysis results
    """
    try:
        # Calculate hotel sets
        current_missing_hotels = set(current_missing['prop_code'].unique())
        previous_missing_hotels = set(previous_missing['prop_code'].unique())
        
        # Identify different hotel categories
        recovered = list(previous_missing_hotels - current_missing_hotels)
        new_missing = list(current_missing_hotels - previous_missing_hotels)
        need_assistance = list(current_missing_hotels & previous_missing_hotels)
        
        # Create comparison datasets
        return ComparisonData(
            recovered_hotels=previous_missing[previous_missing['prop_code'].isin(recovered)][['prop_code', 'hotel']].drop_duplicates(),
            new_missing_hotels=current_missing[current_missing['prop_code'].isin(new_missing)][['prop_code', 'hotel']].drop_duplicates(),
            need_assistance_hotels=current_missing[current_missing['prop_code'].isin(need_assistance)][['prop_code', 'hotel']].drop_duplicates()
        )
    except Exception as e:
        raise Exception(f"Error generating comparison data: {str(e)}")

def generate_insights(missing_df: pd.DataFrame, received_df: pd.DataFrame) -> InsightResult:
    """
    Generate insights from the processed data
    
    Args:
        missing_df: DataFrame containing missing files data
        received_df: DataFrame containing received files data
    
    Returns:
        InsightResult object containing analysis results
    """
    try:
        # Calculate basic metrics
        total_missing_hotels = missing_df['prop_code'].nunique()
        total_received_hotels = received_df['prop_code'].nunique()
        total_missing_files = len(missing_df)
        total_received_files = len(received_df)
        
        # Calculate compliance rate
        total_files = total_missing_files + total_received_files
        compliance_rate = (total_received_files / total_files * 100) if total_files > 0 else 0
        
        # Analyze missing files
        top_missing_files = missing_df['file_name'].value_counts().head(10).to_dict()
        
        # Analyze days missing distribution
        valid_days = pd.to_numeric(missing_df['days_missing'], errors='coerce').dropna()
        days_missing_distribution = {}
        if not valid_days.empty:
            try:
                bins = [0, 7, 14, 30, float('inf')]
                labels = ['1-7 days', '8-14 days', '15-30 days', '30+ days']
                days_missing_bins = pd.cut(valid_days, bins=bins, labels=labels)
                days_missing_distribution = days_missing_bins.value_counts().to_dict()
            except Exception:
                pass
        
        # Calculate additional metrics
        avg_days_missing = valid_days.mean()
        files_per_hotel = missing_df.groupby('prop_code').size().describe().to_dict()
        
        insights_data = {
            'metrics': {
                'total_missing_hotels': total_missing_hotels,
                'total_received_hotels': total_received_hotels,
                'total_missing_files': total_missing_files,
                'total_received_files': total_received_files,
                'compliance_rate': round(compliance_rate, 2),
                'average_days_missing': round(avg_days_missing, 2) if not pd.isna(avg_days_missing) else 0
            },
            'analysis': {
                'top_missing_files': top_missing_files,
                'days_missing_distribution': days_missing_distribution,
                'files_per_hotel_stats': files_per_hotel
            },
            'trends': {
                'by_day': missing_df.groupby('dow').size().to_dict(),
                'by_file_type': missing_df['file_name'].value_counts().to_dict()
            }
        }
        
        return InsightResult(True, "Insights generated successfully", insights_data)
    except Exception as e:
        return InsightResult(False, f"Error generating insights: {str(e)}")

def analyze_hotel_performance(hotel_code: str, missing_df: pd.DataFrame, received_df: pd.DataFrame) -> InsightResult:
    """
    Analyze performance metrics for a specific hotel
    
    Args:
        hotel_code: Property code of the hotel
        missing_df: DataFrame containing missing files data
        received_df: DataFrame containing received files data
    
    Returns:
        InsightResult object containing hotel-specific analysis
    """
    try:
        # Get hotel data
        hotel_missing = missing_df[missing_df['prop_code'] == hotel_code]
        hotel_received = received_df[received_df['prop_code'] == hotel_code]
        
        if hotel_missing.empty and hotel_received.empty:
            return InsightResult(False, f"No data found for hotel {hotel_code}")
        
        # Calculate metrics
        total_missing = len(hotel_missing)
        total_received = len(hotel_received)
        compliance_rate = (total_received / (total_missing + total_received) * 100) if (total_missing + total_received) > 0 else 0
        
        hotel_analysis = {
            'metrics': {
                'total_missing_files': total_missing,
                'total_received_files': total_received,
                'compliance_rate': round(compliance_rate, 2)
            },
            'missing_files': {
                'by_type': hotel_missing['file_name'].value_counts().to_dict(),
                'by_day': hotel_missing.groupby('dow').size().to_dict(),
                'days_missing_stats': hotel_missing['days_missing'].describe().to_dict()
            }
        }
        
        return InsightResult(True, f"Analysis completed for hotel {hotel_code}", hotel_analysis)
    except Exception as e:
        return InsightResult(False, f"Error analyzing hotel {hotel_code}: {str(e)}")

def generate_trend_analysis(current_data: ProcessedData, previous_data: Optional[ProcessedData] = None) -> InsightResult:
    """
    Generate trend analysis comparing current and previous data
    
    Args:
        current_data: Current ProcessedData object
        previous_data: Optional previous ProcessedData object
    
    Returns:
        InsightResult object containing trend analysis
    """
    try:
        current_analysis = {
            'total_missing': len(current_data.missing),
            'unique_hotels': current_data.missing['prop_code'].nunique(),
            'avg_days_missing': current_data.missing['days_missing'].mean(),
            'file_types': current_data.missing['file_name'].value_counts().to_dict()
        }
        
        trend_data = {
            'current': current_analysis,
            'comparison': None
        }
        
        if previous_data:
            previous_analysis = {
                'total_missing': len(previous_data.missing),
                'unique_hotels': previous_data.missing['prop_code'].nunique(),
                'avg_days_missing': previous_data.missing['days_missing'].mean(),
                'file_types': previous_data.missing['file_name'].value_counts().to_dict()
            }
            
            # Calculate changes
            trend_data['comparison'] = {
                'total_missing_change': current_analysis['total_missing'] - previous_analysis['total_missing'],
                'hotels_change': current_analysis['unique_hotels'] - previous_analysis['unique_hotels'],
                'avg_days_change': current_analysis['avg_days_missing'] - previous_analysis['avg_days_missing']
            }
        
        return InsightResult(True, "Trend analysis completed successfully", trend_data)
    except Exception as e:
        return InsightResult(False, f"Error generating trend analysis: {str(e)}")
