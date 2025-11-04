from django.contrib import admin
from django import forms
from django.utils.html import format_html
from django.urls import reverse
from core.models import Person, Conflict, FinancialReport, CreditCard 

class ConflictForm(forms.ModelForm):
    class Meta:
        model = Conflict
        fields = '__all__'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Replace boolean field widgets with custom display
        for field_name in ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11']:
            self.fields[field_name].widget = forms.Select(choices=[(True, 'YES'), (False, 'NO')])
        # Make detail fields optional as they might not always be filled
        for field_name in ['q1_detalle', 'q2_detalle', 'q3_detalle', 'q4_detalle', 'q5_detalle',
                           'q6_detalle', 'q7_detalle', 'q10_detalle', 'q11_detalle']: # New detail fields
            if field_name in self.fields: # Check if field exists to prevent errors if not added in models
                self.fields[field_name].required = False # New field


@admin.register(Person)
class PersonAdmin(admin.ModelAdmin):
    list_display = ('cedula', 'nombre_completo', 'cargo', 'area', 'compania', 'estado', 'revisar')
    search_fields = ('cedula', 'nombre_completo', 'correo')
    list_filter = ('estado', 'compania', 'revisar')
    list_editable = ('revisar',)

    # Custom fields to show in detail view
    readonly_fields = ('cedula_with_actions', 'conflicts_link', 'financial_reports_link')

    fieldsets = (
        (None, {
            'fields': ('cedula_with_actions', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'area', 'revisar', 'comments')
        }),
        ('Related Records', {
            'fields': ('conflicts_link', 'financial_reports_link'),
            'classes': ('collapse',)
        }),
    )

    def cedula_with_actions(self, obj):
        if obj.pk:
            change_url = reverse('admin:core_person_change', args=[obj.pk])
            history_url = reverse('admin:core_person_history', args=[obj.pk])
            add_url = reverse('admin:core_person_add')

            return format_html(
                '{} <div class="nowrap">'
                '<a href="{}" class="changelink">Change</a> &nbsp;'
                '<a href="{}" class="historylink">History</a> &nbsp;'
                '<a href="{}" class="addlink">Add another</a>'
                '</div>',
                obj.cedula,
                change_url,
                history_url,
                add_url
            )
        return obj.cedula
    cedula_with_actions.short_description = 'Cedula'

    def conflicts_link(self, obj):
        if obj.pk:
            conflict = obj.conflicts.first()
            if conflict:
                change_url = reverse('admin:core_conflict_change', args=[conflict.pk])
                add_url = reverse('admin:core_conflict_add') + f'?person={obj.pk}'
                list_url = reverse('admin:core_conflict_changelist') + f'?q={obj.cedula}'

                return format_html(
                    '<div class="nowrap">'
                    '<a href="{}" class="changelink">View/Edit Conflicts</a> &nbsp;'
                    '<a href="{}" class="addlink">Add New Conflict</a> &nbsp;'
                    '<a href="{}" class="viewlink">All Conflicts</a>'
                    '</div>',
                    change_url,
                    add_url,
                    list_url
                )
            else:
                add_url = reverse('admin:core_conflict_add') + f'?person={obj.pk}'
                return format_html(
                    '<a href="{}" class="addlink">Create Conflict Record</a>',
                    add_url
                )
        return "-"
    conflicts_link.short_description = 'Conflict Records'
    conflicts_link.allow_tags = True

    def financial_reports_link(self, obj):
        if obj.pk:
            report = obj.financial_reports.first()
            if report:
                change_url = reverse('admin:core_financialreport_change', args=[report.pk])
                add_url = reverse('admin:core_financialreport_add') + f'?person={obj.pk}'
                list_url = reverse('admin:core_financialreport_changelist') + f'?q={obj.cedula}'

                return format_html(
                    '<div class="nowrap">'
                    '<a href="{}" class="changelink">View/Editar Declaracion B&R</a> &nbsp;'
                    '<a href="{}" class="addlink">Agregar Nueva declaracion B&R</a> &nbsp;'
                    '<a href="{}" class="viewlink">Todo en Bienes y Rentas</a>'
                    '</div>',
                    change_url,
                    add_url,
                    list_url
                )
            else:
                add_url = reverse('admin:core_financialreport_add') + f'?person={obj.pk}'
                return format_html(
                    '<a href="{}" class="addlink">Create Financial Report Record</a>',
                    add_url
                )
        return "-"
    financial_reports_link.short_description = 'Bienes y Rentas'
    financial_reports_link.allow_tags = True

    def get_fieldsets(self, request, obj=None):
        if obj is None:  # Add view
            return [(None, {'fields': ('cedula', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'revisar', 'comments')})]
        return super().get_fieldsets(request, obj)

@admin.register(Conflict)
class ConflictAdmin(admin.ModelAdmin):
    form = ConflictForm
    # Add new detail fields to list_display
    list_display = ('person', 'fecha_inicio',
                    'get_q1_display', 'q1_detalle',
                    'get_q2_display', 'q2_detalle',
                    'get_q3_display', 'q3_detalle',
                    'get_q4_display', 'q4_detalle',
                    'get_q5_display', 'q5_detalle',
                    'get_q6_display', 'q6_detalle',
                    'get_q7_display', 'q7_detalle',
                    'get_q8_display', 'get_q9_display',
                    'get_q10_display', 'q10_detalle',
                    'get_q11_display', 'q11_detalle')
    # Add new detail fields to list_filter and search_fields if desired
    list_filter = ('q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    search_fields = ('person__nombre_completo', 'person__cedula',
                     'q1_detalle', 'q2_detalle', 'q3_detalle', 'q4_detalle', # New search fields
                     'q5_detalle', 'q6_detalle', 'q7_detalle', 'q10_detalle', 'q11_detalle') # New search fields
    raw_id_fields = ('person',)

    # Update fieldsets to include new detail fields
    fieldsets = (
        (None, {
            'fields': ('person', 'fecha_inicio')
        }),
        ('Conflict Questions', {
            'fields': (
                ('q1', 'q1_detalle'), # Group boolean with its detail
                ('q2', 'q2_detalle'), # Group boolean with its detail
                ('q3', 'q3_detalle'), # Group boolean with its detail
                ('q4', 'q4_detalle'), # Group boolean with its detail
                ('q5', 'q5_detalle'), # Group boolean with its detail
                ('q6', 'q6_detalle'), # Group boolean with its detail
                ('q7', 'q7_detalle'), # Group boolean with its detail
                'q8', 'q9',
                ('q10', 'q10_detalle'), # Group boolean with its detail
                ('q11', 'q11_detalle'), # Group boolean with its detail
            ),
            'description': 'Answer "YES" or "NO" to each question and provide details where applicable'
        }),
    )

    def get_form(self, request, obj=None, **kwargs):
        form = super().get_form(request, obj, **kwargs)
        form.base_fields['q1'].label = 'Accionista de proveedor'
        form.base_fields['q1_detalle'].label = 'Accionista de proveedor (Detalle)' # New label
        form.base_fields['q2'].label = 'Familiar de accionista/empleado'
        form.base_fields['q2_detalle'].label = 'Familiar de accionista/empleado (Detalle)' # New label
        form.base_fields['q3'].label = 'Accionista del grupo'
        form.base_fields['q3_detalle'].label = 'Accionista del grupo (Detalle)' # New label
        form.base_fields['q4'].label = 'Actividades extralaborales'
        form.base_fields['q4_detalle'].label = 'Actividades extralaborales (Detalle)' # New label
        form.base_fields['q5'].label = 'Negocios con empleados'
        form.base_fields['q5_detalle'].label = 'Negocios con empleados (Detalle)' # New label
        form.base_fields['q6'].label = 'Participacion en juntas'
        form.base_fields['q6_detalle'].label = 'Participacion en juntas (Detalle)' # New label
        form.base_fields['q7'].label = 'Otro conflicto'
        form.base_fields['q7_detalle'].label = 'Otro conflicto (Detalle)' # New label
        form.base_fields['q8'].label = 'Conoce codigo de conducta'
        form.base_fields['q9'].label = 'Veracidad de informacion'
        form.base_fields['q10'].label = 'Familiar de funcionario'
        form.base_fields['q10_detalle'].label = 'Familiar de funcionario (Detalle)' # New label
        form.base_fields['q11'].label = 'Relacion con sector publico'
        form.base_fields['q11_detalle'].label = 'Relacion con sector publico (Detalle)' # New label
        return form

    # YES/NO display methods for list view (no changes here for detail fields)
    def get_q1_display(self, obj): return "YES" if obj.q1 else "NO"
    get_q1_display.short_description = 'Accionista de proveedor'
    def get_q2_display(self, obj): return "YES" if obj.q2 else "NO"
    get_q2_display.short_description = 'Familiar de accionista/empleado'
    def get_q3_display(self, obj): return "YES" if obj.q3 else "NO"
    get_q3_display.short_description = 'Accionista del grupo'
    def get_q4_display(self, obj): return "YES" if obj.q4 else "NO"
    get_q4_display.short_description = 'Actividades extralaborales'
    def get_q5_display(self, obj): return "YES" if obj.q5 else "NO"
    get_q5_display.short_description = 'Negocios con empleados'
    def get_q6_display(self, obj): return "YES" if obj.q6 else "NO"
    get_q6_display.short_description = 'Participacion en juntas'
    def get_q7_display(self, obj): return "YES" if obj.q7 else "NO"
    get_q7_display.short_description = 'Otro conflicto'
    def get_q8_display(self, obj): return "YES" if obj.q8 else "NO"
    get_q8_display.short_description = 'Conoce codigo de conducta'
    def get_q9_display(self, obj): return "YES" if obj.q9 else "NO"
    get_q9_display.short_description = 'Veracidad de informacion'
    def get_q10_display(self, obj): return "YES" if obj.q10 else "NO"
    get_q10_display.short_description = 'Familiar de funcionario'
    def get_q11_display(self, obj): return "YES" if obj.q11 else "NO"
    get_q11_display.short_description = 'Relacion con sector publico'

@admin.register(FinancialReport) # Register the new model
class FinancialReportAdmin(admin.ModelAdmin):
    list_display = (
        'person', 'fk_id_periodo', 'ano_declaracion', 'activos', 'pasivos',
        'patrimonio', 'ingresos', 'apalancamiento', 'endeudamiento',
        'activos_var_rel', 'pasivos_var_rel', 'patrimonio_var_rel',
        'ingresos_var_rel'
    )
    search_fields = (
        'person__nombre_completo', 'person__cedula', 'fk_id_periodo',
        'ano_declaracion'
    )
    list_filter = ('ano_declaracion', 'fk_id_periodo')
    raw_id_fields = ('person',)

    fieldsets = (
        (None, {
            'fields': ('person', 'fk_id_periodo', 'ano_declaracion', 'ano_creacion')
        }),
        ('Financial Data', {
            'fields': (
                'activos', 'cant_bienes', 'cant_bancos', 'cant_cuentas', 'cant_inversiones',
                'pasivos', 'cant_deudas', 'patrimonio', 'capital', 'aum_pat_subito',
                'banco_saldo', 'bienes', 'inversiones', 'ingresos', 'cant_ingresos'
            )
        }),
        ('Trends and Variations', {
            'fields': (
                ('apalancamiento', 'apalancamiento_var_abs', 'apalancamiento_var_rel'),
                ('endeudamiento', 'endeudamiento_var_abs', 'endeudamiento_var_rel'),
                ('activos_var_abs', 'activos_var_rel'),
                ('pasivos_var_abs', 'pasivos_var_rel'),
                ('patrimonio_var_abs', 'patrimonio_var_rel'),
                ('banco_saldo_var_abs', 'banco_saldo_var_rel'),
                ('bienes_var_abs', 'bienes_var_rel'),
                ('inversiones_var_abs', 'inversiones_var_rel'),
                ('ingresos_var_abs', 'ingresos_var_rel'),
            )
        }),
    )

# ... (Código anterior de admin.py)

# Reemplaza la clase CreditCardAdmin con esta:
@admin.register(CreditCard)
class CreditCardAdmin(admin.ModelAdmin):
    # list_display ahora usa el nombre correcto del campo: 'archivo_nombre'
    list_display = (
        'person_link', 'tipo_tarjeta', 'numero_tarjeta', 'fecha_transaccion', 
        'descripcion', 'valor_cop', 'categoria', 'subcategoria', 'archivo_nombre' 
    )
    
    # search_fields también debe usar el nombre correcto
    search_fields = (
        'person__cedula', 'person__nombre_completo', 'numero_tarjeta', 
        'descripcion', 'categoria', 'subcategoria', 'archivo_nombre' 
    )
    
    # Filtros laterales
    list_filter = (
        'tipo_tarjeta', 'moneda', 'categoria', 'subcategoria', 
        'zona', 'person__compania', 'person__cargo', 'fecha_transaccion'
    )
    
    # ... (El resto del código de la clase CreditCardAdmin, incluyendo person_link y raw_id_fields, es correcto)
    
    def person_link(self, obj):
        link = reverse("admin:core_person_change", args=[obj.person.cedula])
        return format_html('<a href="{}">{} ({})</a>', link, obj.person.nombre_completo, obj.person.cedula)

    person_link.short_description = 'Persona'
    person_link.admin_order_field = 'person__nombre_completo'

    raw_id_fields = ('person',)
