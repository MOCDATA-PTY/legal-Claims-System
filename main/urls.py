from django.urls import path
from . import views

urlpatterns = [
    # =============================================================================
    # AUTHENTICATION & NAVIGATION URLS
    # =============================================================================
    path('', views.index, name='index'),
    path('login/', views.user_login, name='login'),
    path('hidden-register/', views.register, name='register'),
    path('logout/', views.user_logout, name='logout'),
    path('home/', views.home, name='home'),
    
    # =============================================================================
    # CLIENT MANAGEMENT URLS
    # =============================================================================
    path('clients/', views.client_list, name='client_list'),
    path('clients/add/', views.add_client, name='add_client'),
    path('clients/edit/<int:pk>/', views.edit_client, name='edit_client'),
    path('clients/delete/<int:pk>/', views.delete_client, name='delete_client'),
    path('api/client-autocomplete/', views.client_autocomplete, name='client_autocomplete'),
    
    # =============================================================================
    # SHIPMENT MANAGEMENT URLS
    # =============================================================================
    path('shipments/', views.shipment_list, name='shipment_list'),
    path('shipments/add/', views.add_shipment, name='add_shipment'),
    path('shipments/edit/<int:pk>/', views.edit_shipment, name='edit_shipment'),
    path('shipments/delete/<int:pk>/', views.delete_shipment, name='delete_shipment'),
    path('shipments/clear-database/', views.clear_database, name='clear_database'),
    
    # =============================================================================
    # IMPORT/EXPORT URLS
    # =============================================================================
    path('shipments/export/', views.export_shipments, name='export_shipments'),
    path('shipments/export-excel/', views.export_shipments_excel, name='export_shipments_excel'),  # Legacy
    path('shipments/import/', views.import_shipments, name='import_shipments'),
    
    # =============================================================================
    # BACKUP MANAGEMENT URLS
    # =============================================================================
    path('backups/', views.browse_backups, name='browse_backups'),
    path('backups/download/<str:format_type>/<str:filename>/', views.download_backup, name='download_backup'),
    path('backups/manual-backup/', views.manual_backup_now, name='manual_backup_now'),
    path('backups/weekly-status/', views.weekly_backup_status, name='weekly_backup_status'),  # NEW URL
    path('analytics/', views.analytics_dashboard, name='analytics_dashboard'),
]