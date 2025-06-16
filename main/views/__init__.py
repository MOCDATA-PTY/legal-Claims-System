# Import all views from submodules for easy access
from .core_views import (
    # Authentication views
    user_login,
    register,
    user_logout,
    
    # Navigation views
    index,
    home,
    
    # Client management views
    client_list,
    add_client,
    edit_client,
    delete_client,
    client_autocomplete,
    
    # Shipment management views
    add_shipment,
    shipment_list,
    edit_shipment,
    delete_shipment,
    clear_database,
    
    # Analytics views
    analytics_dashboard,
    
    # Helper functions
    apply_filters,
    clear_messages,
)

from .data_views import (
    # Export views
    export_shipments,
    export_shipments_excel,  # Legacy function
    
    # Import views
    import_shipments,
    
    # Backup management views
    browse_backups,
    download_backup,
    manual_backup_now,
    weekly_backup_status,  # NEW VIEW FUNCTION
    
    # Export helper functions
    export_to_excel,
    export_to_csv,
    export_to_pdf,
    process_excel_data,
    
    # Backup helper functions
    setup_backup_directory,
    format_file_size,
    start_backup_thread,
    custom_404
)

# Make all views available at package level
__all__ = [
    # Authentication views
    'user_login',
    'register', 
    'user_logout',
    
    # Navigation views
    'index',
    'home',
    
    # Client management views
    'client_list',
    'add_client',
    'edit_client',
    'delete_client',
    'client_autocomplete',
    
    # Shipment management views
    'add_shipment',
    'shipment_list',
    'edit_shipment',
    'delete_shipment',
    'clear_database',
    
    # Analytics views
    'analytics_dashboard',
    
    # Export/Import views
    'export_shipments',
    'export_shipments_excel',
    'import_shipments',
    
    # Backup management views
    'browse_backups',
    'download_backup',
    'manual_backup_now',
    'weekly_backup_status',  # NEW VIEW FUNCTION
    
    # Helper functions
    'apply_filters',
    'clear_messages',
    'export_to_excel',
    'export_to_csv', 
    'export_to_pdf',
    'process_excel_data',
    'setup_backup_directory',
    'format_file_size',
    'start_backup_thread',
]