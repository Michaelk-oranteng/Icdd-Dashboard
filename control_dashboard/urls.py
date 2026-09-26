# control_dashboard/urls.py

from django.urls import path
from . import views

app_name = 'control_dashboard'

urlpatterns = [
    # ==================== AUTHENTICATION ====================
    path('', views.landing_page, name='landing_page'),
    path('login/', views.landing_page, name='login'),
    path('logout/', views.logout_view, name='logout'),

    # ==================== ADMIN URLS ====================
    path('admin/', views.admin_page, name='admin_page'),

    # ==================== REPORT MANAGEMENT ====================
    path('report-creation/', views.report_creation, name='report_creation'),
    path('report-center/', views.report_center, name='report_center'),

    # ==================== ADMIN ACTIVITY LOGS ====================
    path('activity-logs/', views.activity_logs, name='activity_logs'),

    # ==================== MEMBER URLS ====================
    path('member/', views.member_dashboard, name='member_dashboard'),
    path('member/drafts/', views.drafts_page, name='drafts_page'),
    path('member/reports/', views.reports_page, name='reports_page'),
    path('member/checklist/', views.member_checklist, name='member_checklist'),
    path('member/activity-logs/', views.member_activity_logs, name='member_activity_logs'),

    # ==================== SUPERVISOR URLS ====================
    path('supervisor/', views.supervisor_dashboard, name='supervisor_dashboard'),
    path('supervisor/team-performance/', views.team_performance, name='team_performance'),
    path('supervisor/submitted-reports/', views.submitted_reports, name='submitted_reports'),
    path('supervisor/ad-hoc-scorecard/', views.ad_hoc_scorecard, name='ad_hoc_scorecard'),
    path('supervisor/logged-exceptions/', views.logged_exceptions, name='logged_exceptions'),
    path('supervisor/checklist/', views.supervisor_checklist, name='supervisor_checklist'),
    path('supervisor/activity-logs/', views.supervisor_activity_logs, name='supervisor_activity_logs'),
    path('api/ad-hoc/create/', views.api_create_ad_hoc_deduction, name='api_create_ad_hoc_deduction'),
    path('api/ad-hoc/<int:deduction_id>/update/', views.api_update_ad_hoc_deduction, name='api_update_ad_hoc_deduction'),
    path('api/ad-hoc/<int:deduction_id>/delete/', views.api_delete_ad_hoc_deduction, name='api_delete_ad_hoc_deduction'),

    # ==================== CHECKLIST BUILDER VIEWS ====================
    path('checklists/', views.checklist_builder, name='checklist_builder'),
    path('checklist-list/', views.checklist_list, name='checklist_list'),

    # ==================== API - AUTHENTICATION ====================
    path('api/email-login/', views.api_email_login, name='api_email_login'),
    path('api/user/profile/', views.get_user_profile_api, name='get_user_profile_api'),

    # ==================== API - ADMIN USER MANAGEMENT ====================
    path('api/users/create/', views.api_create_user, name='api_create_user'),
    path('api/users/<int:user_id>/edit/', views.api_edit_user, name='api_edit_user'),
    path('api/users/<int:user_id>/status/', views.api_update_status, name='api_update_status'),
    path('api/users/<int:user_id>/delete/', views.api_delete_user, name='api_delete_user'),

    # ==================== API - REPORT MANAGEMENT ====================
    path('api/reports/create/', views.api_create_report, name='api_create_report'),
    path('api/reports/<int:report_id>/', views.api_get_report, name='api_get_report'),
    path('api/reports/<int:report_id>/edit/', views.api_edit_report, name='api_edit_report'),
    path('api/reports/<int:report_id>/delete/', views.api_delete_report, name='api_delete_report'),

    # ==================== API - MEMBER DRAFTS ====================
    path('api/drafts/<str:report_id>/',         views.api_get_draft,    name='draft_detail'),
    path('api/drafts/<str:report_id>/edit/',    views.api_edit_draft,   name='draft_edit'),
    path('api/drafts/<str:report_id>/delete/',  views.api_delete_draft, name='draft_delete'),

    # ==================== API - CHECKLISTS ====================
    path('api/checklists/create/', views.api_create_checklist, name='api_create_checklist'),
    path('api/checklists/<int:checklist_id>/', views.api_get_checklist, name='api_get_checklist'),
    path('api/checklists/<int:checklist_id>/edit/', views.api_edit_checklist, name='api_edit_checklist'),
    path('api/checklists/<int:checklist_id>/delete/', views.api_delete_checklist, name='api_delete_checklist'),

    # ==================== API - CHECKLIST LOGS ====================
    path('api/checklist-log/', views.api_log_checklist, name='api_log_checklist'),
    path('api/checklist-logs/', views.api_get_checklist_logs, name='api_get_checklist_logs'),
    path('api/checklist-stats/', views.api_get_checklist_stats, name='api_get_checklist_stats'),
    path('api/checklist/detail/<int:user_id>/', views.api_checklist_detail, name='api_checklist_detail'),

    # ==================== API - SUBMIT REPORTS ====================
    path('api/save-draft/', views.api_save_draft, name='api_save_draft'),
    path('api/report-data/', views.api_get_report_data, name='api_get_report_data'),

    # ==================== API - SUBMIT SELECTED REPORTS ====================
    # Points to the working view that redirects to submit_page.
    path('api/submit/selected/', views.submit_selected_reports, name='submit_selected'),

    # ==================== API - EXPORT LOGS ====================
    path('api/export-logs/', views.api_export_logs, name='export_logs'),

    # ==================== EXCEL IMPORT ====================
    path('api/import-excel/', views.api_import_excel, name='api_import_excel'),
    path('api/save-imported-data/', views.api_save_imported_data, name='api_save_imported_data'),

    # ==================== ANALYTICS DASHBOARD ====================
    path('analytics/', views.analytics_dashboard, name='analytics_dashboard'),

    # ==================== DAILY TRIAL BALANCE ====================
    path('member/daily-trial-balance/', views.daily_trial_balance, name='daily_trial_balance'),
    path('api/trial-balance/parse/', views.api_parse_trial_balance, name='api_parse_trial_balance'),
    path('api/trial-balance/submit/', views.api_submit_trial_balance, name='api_submit_trial_balance'),
    path('api/trial-balance/template/', views.api_download_trial_balance_template, name='api_download_trial_balance_template'),
    path('api/trial-balance/by-date/', views.api_get_trial_balance_by_date, name='api_get_trial_balance_by_date'),
    path('api/trial-balance/check-day/', views.api_check_trial_balance_day, name='api_check_trial_balance_day'),

    # ==================== CONSOLIDATED REPORTS ====================
    path('member/consolidated-reports/', views.consolidated_reports, name='consolidated_reports'),
    path('member/consolidated-reports/generate-excel/', views.generate_consolidated_excel, name='generate_consolidated_excel'),
    path('api/consolidated/filter-options/', views.api_consolidated_filter_options, name='api_consolidated_filter_options'),

    # ==================== API - EXCEL ROWS ====================
    path('api/excel-rows/<int:row_id>/delete/', views.api_delete_excel_row, name='excel_row_delete'),
    path('api/excel-rows/<int:row_id>/edit/',   views.api_edit_excel_row,   name='excel_row_edit'),
    path('api/team-performance-live/', views.api_team_performance_live, name='api_team_performance_live'),
    path('api/supervisor-top-performers-live/',views.api_supervisor_top_performers_live,name='api_supervisor_top_performers_live'),
]