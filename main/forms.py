from django import forms
from django.core.exceptions import ValidationError
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django.contrib.auth.models import User
from .models import Shipment, Client

class ClientForm(forms.ModelForm):
    name = forms.CharField(
        required=True,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Client name'
        })
    )
    
    class Meta:
        model = Client
        fields = ['name']  # client_id is auto-generated
        
    def clean_name(self):
        name = self.cleaned_data.get('name')
        if not name:
            raise ValidationError("Client name is required.")
        return name

class ShipmentForm(forms.ModelForm):
    # Add a field for entering client name directly
    client_name = forms.CharField(
        required=True,
        label="Claimant",
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'autocomplete': 'off',
            'list': 'clients-list',
            'placeholder': 'Select or enter claimant name'
        })
    )
    
    Claimed_Amount = forms.DecimalField(
        required=False,
        widget=forms.NumberInput(attrs={'step': '0.01', 'class': 'form-control'})
    )

    class Meta:
        model = Shipment
        fields = [
            'Claim_No', 'client_name', 'Branch', 'Formal_Claim_Received', 
            'Intend_Claim_Date', 'Formal_Claim_Date_Received', 'Claimed_Amount',
            'Amount_Paid_By_Carrier', 'Amount_Paid_By_Awa', 
            'Amount_Paid_By_Insurance', 'Closed_Date'
        ]  # Include client_name at the top
        widgets = {
            'Claim_No': forms.TextInput(attrs={'class': 'form-control', 'required': 'required'}),
            'Branch': forms.Select(attrs={'class': 'form-control', 'required': 'required'}),
            'Intend_Claim_Date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control', 'required': 'required'}),
            'Formal_Claim_Date_Received': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'Closed_Date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'Amount_Paid_By_Carrier': forms.NumberInput(attrs={'step': '0.01', 'class': 'form-control'}),
            'Amount_Paid_By_Awa': forms.NumberInput(attrs={'step': '0.01', 'class': 'form-control'}),
            'Amount_Paid_By_Insurance': forms.NumberInput(attrs={'step': '0.01', 'class': 'form-control'}),
            'Formal_Claim_Received': forms.Select(attrs={'class': 'form-control'}),
        }
        field_order = ['Claim_No', 'client_name', 'Branch', 'Formal_Claim_Received', 
                      'Intend_Claim_Date', 'Formal_Claim_Date_Received', 'Claimed_Amount', 
                      'Amount_Paid_By_Carrier', 'Amount_Paid_By_Awa', 'Amount_Paid_By_Insurance', 
                      'Closed_Date']

    def __init__(self, *args, **kwargs):
        self.instance = kwargs.get('instance')
        super(ShipmentForm, self).__init__(*args, **kwargs)
        
        # If we're editing an existing instance, populate the client_name field
        if self.instance and self.instance.pk and hasattr(self.instance, 'client') and self.instance.client:
            self.fields['client_name'].initial = self.instance.client.name

    def clean_Claim_No(self):
        Claim_No = self.cleaned_data.get('Claim_No')
        if not Claim_No.startswith('S'):
            raise ValidationError("Claim number must start with 'S'.")
        if Shipment.objects.filter(Claim_No=Claim_No).exclude(id=self.instance.id if self.instance else None).exists():
            raise ValidationError("This claim number already exists. Please enter a unique number.")
        return Claim_No

    def clean(self):
        cleaned_data = super().clean()
        # Ensure required fields are provided
        if not cleaned_data.get('client_name'):
            self.add_error('client_name', 'This field is required.')
        if not cleaned_data.get('Branch'):
            self.add_error('Branch', 'This field is required.')
        if not cleaned_data.get('Intend_Claim_Date'):
            self.add_error('Intend_Claim_Date', 'This field is required.')
        return cleaned_data
    
    def save(self, commit=True):
        # Don't save the form yet
        instance = super(ShipmentForm, self).save(commit=False)
        
        # Get or create the client based on the client_name
        client_name = self.cleaned_data.get('client_name')
        if client_name:
            client, created = Client.objects.get_or_create(
                name__iexact=client_name,
                defaults={'name': client_name}
            )
            instance.client = client
        
        # Now save the instance if commit is True
        if commit:
            instance.save()
        
        return instance

# User management forms

class RegisterForm(UserCreationForm):
    email = forms.EmailField(widget=forms.EmailInput(attrs={'class': 'form-control', 'autocomplete': 'off'}))
    username = forms.CharField(widget=forms.TextInput(attrs={'class': 'form-control', 'autocomplete': 'off'}))
    password1 = forms.CharField(label="Password", widget=forms.PasswordInput(attrs={'class': 'form-control', 'autocomplete': 'new-password'}))
    password2 = forms.CharField(label="Confirm Password", widget=forms.PasswordInput(attrs={'class': 'form-control', 'autocomplete': 'new-password'}))

    class Meta:
        model = User
        fields = ['username', 'email', 'password1', 'password2']

class LoginForm(AuthenticationForm):
    username = forms.CharField(widget=forms.TextInput(attrs={'class': 'form-control', 'autocomplete': 'off'}))
    password = forms.CharField(widget=forms.PasswordInput(attrs={'class': 'form-control', 'autocomplete': 'new-password'}))