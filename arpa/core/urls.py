from django.contrib.auth import views as auth_views
from django.urls import path
from . import views
from django.contrib.auth import get_user_model
from django.contrib import messages
from django.shortcuts import render, redirect
from django.urls import path
from django.contrib.auth import views as auth_views
from .views import (main, register_superuser, ImportView, person_list,
                   import_conflicts, conflict_list, import_persons,
                   import_finances, person_details, financial_report_list,
                   export_persons_excel, alerts_list, save_comment, delete_comment, import_tcs, import_categorias,
                   tcs_list, import_personas_tc, export_financial_reports_excel, export_credit_card_excel) 

def register_superuser(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password1 = request.POST.get('password1')
        password2 = request.POST.get('password2')

        if password1 != password2:
            messages.error(request, "Passwords don't match")
            return redirect('register')

        User = get_user_model()
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return redirect('register')

        try:
            user = User.objects.create_superuser(
                username=username,
                email=email,
                password=password1
            )
            messages.success(request, f"Superuser {username} created successfully!")
            return redirect('login')
        except Exception as e:
            messages.error(request, f"Error creating superuser: {str(e)}")
            return redirect('register')

    return render(request, 'registration/register.html')

urlpatterns = [
    path('', main, name='main'),
    path('login/', auth_views.LoginView.as_view(template_name='registration/login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('register/', register_superuser, name='register'),
    path('import/', ImportView.as_view(), name='import'),
    path('import/persons/', import_persons, name='import_persons'),
    path('import/personas_tc/', import_personas_tc, name='import_personas_tc'),
    path('import/conflicts/', import_conflicts, name='import_conflicts'),
    path('import/finances/', import_finances, name='import_finances'),
    path('import/tcs/', import_tcs, name='import_tcs'),
    path('import/categorias/', import_categorias, name='import_categorias'),
    path('persons/', person_list, name='person_list'),
    path('persons/<str:cedula>/', person_details, name='person_details'),
    path('persons/export/excel/', export_persons_excel, name='export_persons_excel'),
    path('financial_reports/', financial_report_list, name='financial_report_list'),
    path('alerts/', alerts_list, name='alerts_list'),
    path('persons/<str:cedula>/toggle_revisar/', views.toggle_revisar_status, name='toggle_revisar_status'),
    path('persons/<str:cedula>/save_comment/', save_comment, name='save_comment'),
    path('persons/<str:cedula>/delete_comment/<int:comment_index>/', delete_comment, name='delete_comment'),
    path('tcs/', tcs_list, name='tcs_list'), 
    path('conflicts/', conflict_list, name='conflict_list'),
    path('financial_reports/export/excel/', export_financial_reports_excel, name='export_financial_reports_excel'), 
    path('tcs/export/excel/', export_credit_card_excel, name='export_credit_card_excel'),
    
]
