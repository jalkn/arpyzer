from django.db import models
from django.contrib.auth.models import User

class Person(models.Model):
    cedula = models.CharField(max_length=20, primary_key=True)
    nombre_completo = models.CharField(max_length=255)
    correo = models.EmailField(max_length=255, blank=True)
    estado = models.CharField(max_length=50, default='Activo')
    compania = models.CharField(max_length=255, blank=True)
    cargo = models.CharField(max_length=255, blank=True)
    area = models.CharField(max_length=255, blank=True)
    revisar = models.BooleanField(default=False)
    comments = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.nombre_completo} ({self.cedula})"

class Conflict(models.Model):
    id = models.AutoField(primary_key=True)
    person = models.ForeignKey(
        Person,
        on_delete=models.CASCADE,
        related_name='conflicts',
        to_field='cedula',
        db_column='cedula'
    )
    ano = models.IntegerField(null=True, blank=True)
    fecha_inicio = models.DateField(null=True, blank=True)
    q1 = models.BooleanField(null=True, blank=True) 
    q1_detalle = models.TextField(blank=True)
    q2 = models.BooleanField(null=True, blank=True) 
    q2_detalle = models.TextField(blank=True)
    q3 = models.BooleanField(null=True, blank=True) 
    q3_detalle = models.TextField(blank=True)
    q4 = models.BooleanField(null=True, blank=True) 
    q4_detalle = models.TextField(blank=True)
    q5 = models.BooleanField(null=True, blank=True) 
    q5_detalle = models.TextField(blank=True)
    q6 = models.BooleanField(null=True, blank=True) 
    q6_detalle = models.TextField(blank=True)
    q7 = models.BooleanField(null=True, blank=True) 
    q7_detalle = models.TextField(blank=True)
    q8 = models.BooleanField(null=True, blank=True) 
    q9 = models.BooleanField(null=True, blank=True) 
    q10 = models.BooleanField(null=True, blank=True) 
    q10_detalle = models.TextField(blank=True)
    q11 = models.BooleanField(null=True, blank=True) 
    q11_detalle = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Conflictos para {self.person.nombre_completo} (ID: {self.id}, Año: {self.ano})"

class FinancialReport(models.Model):
    id = models.AutoField(primary_key=True)
    person = models.ForeignKey(
        Person,
        on_delete=models.CASCADE,
        related_name='financial_reports',
        to_field='cedula',
        db_column='cedula'
    )
    fk_id_periodo = models.IntegerField(null=True, blank=True)
    ano_declaracion = models.IntegerField(null=True, blank=True)
    ano_creacion = models.IntegerField(null=True, blank=True)
    activos = models.FloatField(null=True, blank=True)
    cant_bienes = models.IntegerField(null=True, blank=True)
    cant_bancos = models.IntegerField(null=True, blank=True)
    cant_cuentas = models.IntegerField(null=True, blank=True)
    cant_inversiones = models.IntegerField(null=True, blank=True)
    pasivos = models.FloatField(null=True, blank=True)
    cant_deudas = models.IntegerField(null=True, blank=True)
    patrimonio = models.FloatField(null=True, blank=True)
    apalancamiento = models.FloatField(null=True, blank=True)
    endeudamiento = models.FloatField(null=True, blank=True)
    capital = models.FloatField(null=True, blank=True)
    aum_pat_subito = models.FloatField(null=True, blank=True)
    activos_var_abs = models.FloatField(null=True, blank=True)
    activos_var_rel = models.CharField(max_length=50, null=True, blank=True)
    pasivos_var_abs = models.FloatField(null=True, blank=True)
    pasivos_var_rel = models.CharField(max_length=50, null=True, blank=True)
    patrimonio_var_abs = models.FloatField(null=True, blank=True)
    patrimonio_var_rel = models.CharField(max_length=50, null=True, blank=True)
    apalancamiento_var_abs = models.FloatField(null=True, blank=True)
    apalancamiento_var_rel = models.CharField(max_length=50, null=True, blank=True)
    endeudamiento_var_abs = models.FloatField(null=True, blank=True)
    endeudamiento_var_rel = models.CharField(max_length=50, null=True, blank=True)
    banco_saldo = models.FloatField(null=True, blank=True)
    bienes = models.FloatField(null=True, blank=True)
    inversiones = models.FloatField(null=True, blank=True)
    banco_saldo_var_abs = models.FloatField(null=True, blank=True)
    banco_saldo_var_rel = models.CharField(max_length=50, null=True, blank=True)
    bienes_var_abs = models.FloatField(null=True, blank=True)
    bienes_var_rel = models.CharField(max_length=50, null=True, blank=True)
    inversiones_var_abs = models.FloatField(null=True, blank=True)
    inversiones_var_rel = models.CharField(max_length=50, null=True, blank=True)
    ingresos = models.FloatField(null=True, blank=True)
    cant_ingresos = models.IntegerField(null=True, blank=True)
    ingresos_var_abs = models.FloatField(null=True, blank=True)
    ingresos_var_rel = models.CharField(max_length=50, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Reporte Financiero para {self.person.nombre_completo} (Periodo: {self.fk_id_periodo})"
    

class CreditCard(models.Model):
    id = models.AutoField(primary_key=True)

    person = models.ForeignKey(
        Person,
        on_delete=models.SET_NULL,
        related_name='credit_cards', 
        to_field='cedula',
        db_column='cedula', 
        null=True,
        blank=True
    )

    tipo_tarjeta = models.CharField(max_length=50, null=True, blank=True)
    numero_tarjeta = models.CharField(max_length=20, null=True, blank=True) # Número de Tarjeta
    moneda = models.CharField(max_length=10, null=True, blank=True)
    trm_cierre = models.CharField(max_length=50, null=True, blank=True) # Renombrado y tipo ajustado a cómo lo genera tcs.py (string)
    valor_original = models.CharField(max_length=50, null=True, blank=True) # Tipo ajustado a string
    valor_cop = models.CharField(max_length=50, null=True, blank=True) # Agregado de nuevo, tipo ajustado a string
    numero_autorizacion = models.CharField(max_length=100, null=True, blank=True)
    fecha_transaccion = models.DateField(null=True, blank=True)
    dia = models.CharField(max_length=20, null=True, blank=True)
    descripcion = models.TextField(null=True, blank=True)
    categoria = models.CharField(max_length=255, null=True, blank=True)
    subcategoria = models.CharField(max_length=255, null=True, blank=True)
    zona = models.CharField(max_length=255, null=True, blank=True)
    
    archivo_nombre = models.CharField(max_length=255, null=True, blank=True) 

    def __str__(self):
        return f"{self.descripcion} - {self.valor_cop} (Tarjeta: {self.numero_tarjeta})"

    class Meta:
        verbose_name = "Tarjeta de Crédito"
        verbose_name_plural = "Tarjetas de Crédito"
        unique_together = ('person', 'fecha_transaccion', 'numero_autorizacion', 'valor_original')
