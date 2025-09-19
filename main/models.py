from django.db import models, transaction
from django.core.cache import cache
from datetime import date
import re

class ClientManager(models.Manager):
    def get_next_client_id(self):
        """Generate next sequential client ID (CL00001, CL00002, etc.)"""
        with transaction.atomic():
            last_client = self.order_by('-id').first()
            if last_client:
                last_id = int(last_client.client_id[2:])
                return f"CL{(last_id + 1):05d}"
            else:
                return "CL00001"

class Client(models.Model):
    client_id = models.CharField(max_length=10, unique=True, db_index=True)
    name = models.CharField(max_length=100, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    objects = ClientManager()
    
    class Meta:
        indexes = [
            models.Index(fields=['client_id']),
            models.Index(fields=['name']),
            models.Index(fields=['-created_at']),
        ]
    
    def __str__(self):
        return f"{self.name} ({self.client_id})"
    
    def save(self, *args, **kwargs):
        if not self.client_id:
            self.client_id = Client.objects.get_next_client_id()
        super().save(*args, **kwargs)
    
    def get_next_client_reference(self):
        """Generate unique client reference for this client: ClientName-X-YYYYMMDD"""
        today = date.today().strftime("%Y%m%d")
        
        # Clean client name (remove spaces, special chars, limit length)
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', self.name)[:15]
        
        with transaction.atomic():
            # Get the highest number for this client today
            existing_shipments = Shipment.objects.filter(
                client=self,
                client_reference__startswith=f"{clean_name}-",
                client_reference__endswith=f"-{today}"
            ).order_by('-id')
            
            next_number = 1
            if existing_shipments.exists():
                # Extract the highest number from existing shipments
                for shipment in existing_shipments:
                    try:
                        # Parse: ClientName-X-YYYYMMDD
                        parts = shipment.client_reference.split('-')
                        if len(parts) >= 3:
                            number = int(parts[-2])  # Get the number part
                            if number >= next_number:
                                next_number = number + 1
                                break
                    except (ValueError, IndexError):
                        continue
            
            return f"{clean_name}-{next_number}-{today}"

class ShipmentManager(models.Manager):
    def get_queryset(self):
        return super().get_queryset().select_related('client')
    
    def get_next_claim_id(self):
        """Generate next sequential claim ID (CLM000001, CLM000002, etc.)"""
        with transaction.atomic():
            last_shipment = self.order_by('-id').first()
            if last_shipment and last_shipment.claim_id:
                # Extract number from existing claim_id
                try:
                    last_id = int(last_shipment.claim_id[3:])  # Remove 'CLM' prefix
                    return f"CLM{(last_id + 1):06d}"
                except (ValueError, IndexError):
                    pass
            return "CLM000001"

class Shipment(models.Model):
    BRANCH_CHOICES = (
        ('ATL', 'ATL'),
        ('CMU', 'CMU'),
        ('CON', 'CON'),
        ('DOR', 'DOR'),
        ('HEC', 'HEC'),
        ('HNL', 'HNL'),
        ('HOU', 'HOU'),
        ('ICS', 'ICS'),
        ('IMP', 'IMP'),
        ('JFK', 'JFK'),
        ('LAX', 'LAX'),
        ('LCL', 'LCL'),
        ('ORD', 'ORD'),
        ('PPG', 'PPG'),
    )
    
    YES_NO_CHOICES = (
        ('YES', 'Yes'),
        ('NO', 'No'),
    )
    
    STATUS_CHOICES = (
        ('OPEN', 'Open'),
        ('PENDING', 'Pending'),
        ('CLOSED', 'Closed'),
        ('REJECTED', 'Rejected'),
        ('UNDER_REVIEW', 'Under Review'),
    )
    
    SETTLEMENT_CHOICES = (
        ('SETTLED', 'Settled'),
        ('NOT_SETTLED', 'Not Settled'),
        ('PARTIAL', 'Partially Settled'),
    )
    
    # KEEP ORIGINAL: Core shipment identification
    Claim_No = models.CharField(max_length=100, unique=True, verbose_name="Shipment Number")
    
    # UNIQUE CLAIM ID: Auto-generated unique claim identifier
    claim_id = models.CharField(
        max_length=20, 
        unique=True, 
        blank=True, 
        null=True,
        verbose_name="Claim ID",
        help_text="Auto-generated unique claim identifier (e.g., CLM000001, CLM000002)"
    )
    
    # NEW FIELD: Client-specific reference (auto-generated)
    client_reference = models.CharField(
        max_length=50, 
        blank=True, 
        null=True, 
        verbose_name="Client Reference",
        help_text="Auto-generated client-specific reference (e.g., ClientName-1-20250601)"
    )
    
    client = models.ForeignKey(Client, on_delete=models.CASCADE, related_name='shipments')
    Branch = models.CharField(max_length=3, choices=BRANCH_CHOICES)
    
    # Brand and Claimant fields
    Brand = models.CharField(max_length=100, blank=True, null=True, verbose_name="Brand")
    Claimant = models.CharField(max_length=200, blank=True, null=True, verbose_name="Claimant Name")
    
    # Intent to claim information
    Intent_To_Claim = models.CharField(max_length=3, choices=YES_NO_CHOICES, blank=True, null=True, verbose_name="Intent To Claim")
    Intend_Claim_Date = models.DateField(blank=True, null=True, verbose_name="Intent To Claim Date")
    
    # Formal claim information
    Formal_Claim_Received = models.CharField(max_length=3, choices=YES_NO_CHOICES, blank=True, null=True, verbose_name="Formal Claim")
    Formal_Claim_Date_Received = models.DateField(blank=True, null=True, verbose_name="Formal Claim Date")
    
    # Financial information
    Claimed_Amount = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Value")
    Amount_Paid_By_Carrier = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Paid By Carrier")
    Amount_Paid_By_Awa = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Paid By ISCM/AWA")
    Amount_Paid_By_Insurance = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Paid By Insurance")
    
    # Additional financial fields
    Total_Savings = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Total Savings")
    Financial_Exposure = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Financial Exposure")
    
    # Settlement and status tracking
    Settlement_Status = models.CharField(max_length=15, choices=SETTLEMENT_CHOICES, blank=True, null=True, verbose_name="Settled or Not Settled")
    Status = models.CharField(max_length=15, choices=STATUS_CHOICES, default='OPEN', verbose_name="Status")
    
    # Date tracking
    Closed_Date = models.DateField(blank=True, null=True, verbose_name="Closed Date")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    objects = ShipmentManager()
    
    class Meta:
        ordering = ['-created_at']
        verbose_name = "Shipment"
        verbose_name_plural = "Shipments"
        indexes = [
            models.Index(fields=['client', 'Claim_No']),
            models.Index(fields=['Claim_No']),
            models.Index(fields=['client_reference']),  # NEW INDEX
            models.Index(fields=['Status']),
            models.Index(fields=['Settlement_Status']),
            models.Index(fields=['-created_at']),
        ]
    
    def __str__(self):
        return f"Shipment No: {self.Claim_No}"
    
    @property
    def total_amount_paid(self):
        """Calculate total amount paid from all sources."""
        carrier = self.Amount_Paid_By_Carrier or 0
        awa = self.Amount_Paid_By_Awa or 0
        insurance = self.Amount_Paid_By_Insurance or 0
        return carrier + awa + insurance
    
    @property
    def outstanding_amount(self):
        """Calculate outstanding amount (claimed minus paid)."""
        claimed = self.Claimed_Amount or 0
        paid = self.total_amount_paid
        return max(0, claimed - paid)
    
    @property
    def is_fully_settled(self):
        """Check if claim is fully settled."""
        return self.Settlement_Status == 'SETTLED'
    
    def save(self, *args, **kwargs):
        # Auto-generate claim_id if not provided
        if not self.claim_id:
            self.claim_id = Shipment.objects.get_next_claim_id()
        
        # Auto-generate client reference if not provided
        if not self.client_reference and self.client:
            self.client_reference = self.client.get_next_client_reference()
        
        # Auto-calculate total savings if not provided
        if self.Total_Savings is None:
            claimed = self.Claimed_Amount or 0
            paid = self.total_amount_paid
            if claimed > 0:
                self.Total_Savings = max(0, claimed - paid)
        
        # Auto-update settlement status based on payments
        if self.Claimed_Amount and self.total_amount_paid:
            if self.total_amount_paid >= self.Claimed_Amount:
                self.Settlement_Status = 'SETTLED'
            elif self.total_amount_paid > 0:
                self.Settlement_Status = 'PARTIAL'
            else:
                self.Settlement_Status = 'NOT_SETTLED'
        
        super().save(*args, **kwargs)