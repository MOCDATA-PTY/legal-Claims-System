from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db import transaction
from ..forms import ShipmentForm, LoginForm, RegisterForm, ClientForm
from ..models import Shipment, Client


def clear_messages(request):
    """Helper function to clear all messages from the request."""
    storage = messages.get_messages(request)
    for message in storage:
        pass
    storage.used = True


def index(request):
    """Redirect root URL to login page."""
    return redirect('login')


# =============================================================================
# AUTHENTICATION VIEWS
# =============================================================================

def user_login(request):
    """Handle user login with improved error handling."""
    if request.method == 'POST':
        form = LoginForm(request, data=request.POST)
        username = request.POST.get('username')
        password = request.POST.get('password')
        
        # Debug print
        print(f"Login attempt - Username: {username}")
        
        user = authenticate(request, username=username, password=password)
        if user is not None:
            if user.is_active:
                login(request, user)
                request.session.set_expiry(0)  # Expire session on browser close
                messages.success(request, f"Welcome back, {user.username}!")
                # Check if there's a next parameter
                next_url = request.GET.get('next') or request.POST.get('next')
                if next_url:
                    return redirect(next_url)
                return redirect('home')
            else:
                messages.error(request, "Your account is disabled.")
        else:
            messages.error(request, "Invalid username or password. Please try again.")
    else:
        form = LoginForm()

    response = render(request, 'main/login.html', {'form': form})
    response['Cache-Control'] = 'no-store'
    return response


def register(request):
    """Handle new user registration and automatic login upon successful registration."""
    clear_messages(request)
    if request.method == 'POST':
        form = RegisterForm(request.POST)
        if form.is_valid():
            try:
                user = form.save()
                # Log the user in immediately after registration
                user = authenticate(request, username=user.username, password=form.cleaned_data['password1'])
                if user:
                    login(request, user)
                    request.session.set_expiry(0)
                    messages.success(request, f"Registration successful! Welcome, {user.username}!")
                    return redirect('home')
                else:
                    messages.error(request, "Registration successful but login failed. Please try logging in manually.")
                    return redirect('login')
            except Exception as e:
                print(f"Registration error: {e}")
                messages.error(request, f"Registration failed: {str(e)}")
        else:
            # Debug: print form errors and add them as messages
            print("Form errors:", form.errors)
            for field, errors in form.errors.items():
                for error in errors:
                    if field == '__all__':
                        messages.error(request, f"{error}")
                    elif field == 'username':
                        messages.error(request, f"Username: {error}")
                    elif field == 'password1':
                        messages.error(request, f"Password: {error}")
                    elif field == 'password2':
                        messages.error(request, f"Password confirmation: {error}")
                    elif field == 'email':
                        messages.error(request, f"Email: {error}")
                    else:
                        messages.error(request, f"{field.replace('_', ' ').title()}: {error}")
    else:
        form = RegisterForm()
    
    return render(request, 'main/register.html', {'form': form})


def user_logout(request):
    """Handle user logout and redirect to login page."""
    clear_messages(request)
    logout(request)  # Clears user session
    messages.success(request, "You have been logged out successfully.")
    return redirect('login')


# =============================================================================
# HOME & NAVIGATION
# =============================================================================

@login_required(login_url='login')
def home(request):
    """Home page view that requires login."""
    clear_messages(request)
    return render(request, 'main/home.html')


# =============================================================================
# CLIENT MANAGEMENT VIEWS
# =============================================================================

@login_required(login_url='login')
def client_list(request):
    """View to list all clients."""
    clear_messages(request)
    clients = Client.objects.all().order_by('name')
    return render(request, 'main/client_list.html', {'clients': clients})


@login_required(login_url='login')
def add_client(request):
    """Add a new client."""
    clear_messages(request)
    if request.method == 'POST':
        form = ClientForm(request.POST)
        if form.is_valid():
            client = form.save()
            messages.success(request, f'Client added successfully with ID: {client.client_id}')
            return redirect('client_list')
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = ClientForm()
    
    return render(request, 'main/add_client.html', {'form': form})


@login_required(login_url='login')
def edit_client(request, pk):
    """Edit an existing client."""
    clear_messages(request)
    client = get_object_or_404(Client, pk=pk)
    
    if request.method == 'POST':
        form = ClientForm(request.POST, instance=client)
        if form.is_valid():
            form.save()
            messages.success(request, 'Client updated successfully.')
            return redirect('client_list')
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = ClientForm(instance=client)
    
    return render(request, 'main/edit_client.html', {'form': form, 'client': client})


@login_required(login_url='login')
def delete_client(request, pk):
    """Delete a client."""
    clear_messages(request)
    client = get_object_or_404(Client, pk=pk)
    
    # Check if there are shipments associated with this client
    shipment_count = Shipment.objects.filter(client=client).count()
    
    if request.method == 'POST':
        if shipment_count > 0 and 'confirm' not in request.POST:
            messages.warning(request, f'This client has {shipment_count} shipments. Are you sure you want to delete?')
            return render(request, 'main/delete_client_confirm.html', {'client': client, 'shipment_count': shipment_count})
        
        # Delete client and any associated shipments
        client.delete()
        messages.success(request, 'Client deleted successfully.')
        return redirect('client_list')
    
    return render(request, 'main/delete_client_confirm.html', {'client': client, 'shipment_count': shipment_count})


@login_required(login_url='login')
def client_autocomplete(request):
    """Autocomplete for client names."""
    term = request.GET.get('term', '')
    clients = Client.objects.filter(name__icontains=term).values('id', 'name', 'client_id')
    
    suggestions = [{'id': client['id'], 'text': f"{client['name']} ({client['client_id']})"} for client in clients]

    return JsonResponse(suggestions, safe=False)


# =============================================================================
# SHIPMENT MANAGEMENT VIEWS
# =============================================================================

@login_required(login_url='login')
def add_shipment(request):
    """Add a new shipment."""
    # Get all clients for dropdown
    clients = Client.objects.all().order_by('name')

    if request.method == 'POST':
        form = ShipmentForm(request.POST)
        claim_no = request.POST.get('Claim_No')

        if Shipment.objects.filter(Claim_No=claim_no).exists():
            existing_shipment = Shipment.objects.get(Claim_No=claim_no)
            messages.warning(request, f'Duplicate claim number {claim_no}, consider editing the existing entry.')
            return render(request, 'main/add_shipment.html', {
                'form': form,
                'clients': clients,
                'duplicate_claim_no': claim_no,
                'edit_shipment_id': existing_shipment.id
            })

        if form.is_valid():
            try:
                # Check if client exists
                client_name = form.cleaned_data.get('client_name')
                client, created = Client.objects.get_or_create(
                    name__iexact=client_name,
                    defaults={'name': client_name}
                )
                
                # Add success message about client
                if created:
                    messages.info(request, f'New client "{client_name}" created with ID: {client.client_id}')
                else:
                    messages.info(request, f'Using existing client "{client_name}" (ID: {client.client_id})')
                
                # Save the form
                shipment = form.save()
                messages.success(request, f'Shipment {shipment.Claim_No} added successfully.')
                return redirect('shipment_list')
            except Exception as e:
                print(f"Error saving shipment: {e}")
                messages.error(request, f'Error saving shipment: {str(e)}')
        else:
            print("Form errors:", form.errors)
            messages.error(request, 'Please correct the errors below.')

    else:
        form = ShipmentForm()

    return render(request, 'main/add_shipment.html', {
        'form': form,
        'clients': clients
    })


@login_required(login_url='login')
def shipment_list(request):
    """List all shipments with the option to apply filters."""
    clear_messages(request)
    shipments = Shipment.objects.all()
    branches = Shipment.objects.values_list('Branch', flat=True).distinct().order_by('Branch')
    clients = Client.objects.all().order_by('name')
    
    # Apply filters
    shipments = apply_filters(request, shipments)
    
    # Handle legacy data - ensure all shipments have clients
    for shipment in shipments:
        if not hasattr(shipment, 'client') or shipment.client is None:
            # Create a client for this shipment if it doesn't have one
            if hasattr(shipment, 'Claiming_Client') and shipment.Claiming_Client:
                client, created = Client.objects.get_or_create(
                    name=shipment.Claiming_Client,
                    defaults={'name': shipment.Claiming_Client}
                )
                # Link the client to the shipment
                shipment.client = client
                shipment.save(update_fields=['client'])
    
    return render(request, 'main/shipment_list.html', {
        'shipments': shipments,
        'branches': branches,
        'clients': clients,
    })


@login_required(login_url='login')
def edit_shipment(request, pk):
    """Edit an existing shipment."""
    clear_messages(request)
    shipment = get_object_or_404(Shipment, pk=pk)
    clients = Client.objects.all().order_by('name')
    
    # Check if this is a cancel/keep original request
    if request.method == 'POST' and 'keep_original' in request.POST:
        messages.info(request, 'No changes were made to the shipment.')
        return redirect('shipment_list')
    
    if request.method == 'POST':
        form = ShipmentForm(request.POST, instance=shipment)
        if form.is_valid():
            try:
                # Check if client exists
                client_name = form.cleaned_data.get('client_name')
                client, created = Client.objects.get_or_create(
                    name__iexact=client_name,
                    defaults={'name': client_name}
                )
                
                # Add success message about client
                if created:
                    messages.info(request, f'New client "{client_name}" created with ID: {client.client_id}')
                else:
                    messages.info(request, f'Using existing client "{client_name}" (ID: {client.client_id})')
                
                # Save the form
                updated_shipment = form.save()
                messages.success(request, f'Shipment {updated_shipment.Claim_No} updated successfully.')
                return redirect('shipment_list')
            except Exception as e:
                print(f"Error updating shipment: {e}")
                messages.error(request, f'Error updating shipment: {str(e)}')
        else:
            print("Form errors:", form.errors)
            messages.error(request, 'Please correct the errors below.')
    else:
        form = ShipmentForm(instance=shipment)
    
    return render(request, 'main/edit_shipment.html', {
        'form': form, 
        'shipment': shipment,
        'clients': clients
    })


@login_required(login_url='login')
def delete_shipment(request, pk):
    """Delete a shipment entry."""
    clear_messages(request)
    shipment = get_object_or_404(Shipment, pk=pk)
    if request.method == 'POST':
        claim_no = shipment.Claim_No
        shipment.delete()
        messages.success(request, f'Shipment {claim_no} deleted successfully.')
    return redirect('shipment_list')


@login_required(login_url='login')
def clear_database(request):
    """Clear all shipments from the database."""
    if request.method == 'POST':
        try:
            with transaction.atomic():
                count = Shipment.objects.count()
                Shipment.objects.all().delete()
                messages.success(request, f"Successfully deleted {count} shipments from the database.")
        except Exception as e:
            messages.error(request, f"Error clearing database: {str(e)}")
        return redirect('shipment_list')
    else:
        messages.error(request, "Invalid request method.")
        return redirect('shipment_list')


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def apply_filters(request, shipments):
    """Apply filters to the shipments queryset based on request parameters."""
    
    # Filter by claim number
    claim_no = request.GET.get('claim_no')
    if claim_no:
        shipments = shipments.filter(Claim_No__icontains=claim_no)
    
    # Filter by client (using the client id or name)
    client_id = request.GET.get('client')
    if client_id:
        try:
            # First try if client_id is numeric (direct ID)
            client_id = int(client_id)
            shipments = shipments.filter(client_id=client_id)
        except ValueError:
            # If not numeric, search by name
            shipments = shipments.filter(client__name__icontains=client_id)
    
    # Filter by client ID (specific Client ID field)
    client_unique_id = request.GET.get('client_unique_id')
    if client_unique_id:
        shipments = shipments.filter(client__client_id=client_unique_id)
    
    # Filter by branch
    branch = request.GET.get('branch')
    if branch:
        shipments = shipments.filter(Branch=branch)
    
    # Filter by Intend Date range
    intend_date_from = request.GET.get('intend_date_from')
    if intend_date_from:
        shipments = shipments.filter(Intend_Claim_Date__gte=intend_date_from)
    
    intend_date_to = request.GET.get('intend_date_to')
    if intend_date_to:
        shipments = shipments.filter(Intend_Claim_Date__lte=intend_date_to)
    
    # Filter by Formal Date range
    formal_date_from = request.GET.get('formal_date_from')
    if formal_date_from:
        shipments = shipments.filter(Formal_Claim_Date_Received__gte=formal_date_from)
    
    formal_date_to = request.GET.get('formal_date_to')
    if formal_date_to:
        shipments = shipments.filter(Formal_Claim_Date_Received__lte=formal_date_to)
    
    return shipments

@login_required(login_url='login')
def analytics_dashboard(request):
    """Data-focused analytics dashboard with comprehensive metrics."""
    clear_messages(request)
    
    from django.db import models
    from django.utils import timezone
    from datetime import timedelta, date
    from decimal import Decimal
    import calendar
    
    try:
        # Get all shipments for analysis
        shipments = Shipment.objects.select_related('client').all()
        
        # === BASIC STATISTICS WITH SAFE AGGREGATION ===
        total_claims = shipments.count()
        
        # Safe aggregation with null handling
        total_value = float(shipments.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0)
        total_paid_iscm = float(shipments.aggregate(total=models.Sum('Amount_Paid_By_Awa'))['total'] or 0)
        total_paid_carrier = float(shipments.aggregate(total=models.Sum('Amount_Paid_By_Carrier'))['total'] or 0)
        total_paid_insurance = float(shipments.aggregate(total=models.Sum('Amount_Paid_By_Insurance'))['total'] or 0)
        total_savings = float(shipments.aggregate(total=models.Sum('Total_Savings'))['total'] or 0)
        total_exposure = float(shipments.aggregate(total=models.Sum('Financial_Exposure'))['total'] or 0)
        
        # === CALCULATED METRICS WITH SAFE DIVISION ===
        total_paid_all = total_paid_iscm + total_paid_carrier + total_paid_insurance
        recovery_rate = (total_paid_all / total_value * 100) if total_value > 0 else 0
        savings_rate = (total_savings / total_value * 100) if total_value > 0 else 0
        exposure_rate = (total_exposure / total_value * 100) if total_value > 0 else 0
        avg_claim_value = total_value / total_claims if total_claims > 0 else 0
        avg_savings_per_claim = total_savings / total_claims if total_claims > 0 else 0
        avg_exposure_per_claim = total_exposure / total_claims if total_claims > 0 else 0
        
        # === SIMPLIFIED TIME-BASED ANALYSIS ===
        current_date = timezone.now().date()
        
        # Last 30 days analysis
        last_30_days = current_date - timedelta(days=30)
        recent_claims = shipments.filter(Intend_Claim_Date__gte=last_30_days)
        
        time_analysis = {
            '30_days': {
                'count': recent_claims.count(),
                'value': float(recent_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                'savings': float(recent_claims.aggregate(total=models.Sum('Total_Savings'))['total'] or 0),
                'avg_value': float(recent_claims.aggregate(avg=models.Avg('Claimed_Amount'))['avg'] or 0),
                'settled_count': recent_claims.filter(Settlement_Status='SETTLED').count(),
                'open_count': recent_claims.filter(Status='OPEN').count(),
            }
        }
        
        # === YEARLY ANALYSIS - SIMPLIFIED ===
        current_year = current_date.year
        current_year_claims = shipments.filter(Intend_Claim_Date__year=current_year)
        last_year_claims = shipments.filter(Intend_Claim_Date__year=current_year-1)
        
        yearly_data = {
            current_year: {
                'count': current_year_claims.count(),
                'value': float(current_year_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                'savings': float(current_year_claims.aggregate(total=models.Sum('Total_Savings'))['total'] or 0),
                'settled_rate': (current_year_claims.filter(Settlement_Status='SETTLED').count() / current_year_claims.count() * 100) if current_year_claims.count() > 0 else 0,
            },
            current_year-1: {
                'count': last_year_claims.count(),
                'value': float(last_year_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                'savings': float(last_year_claims.aggregate(total=models.Sum('Total_Savings'))['total'] or 0),
                'settled_rate': (last_year_claims.filter(Settlement_Status='SETTLED').count() / last_year_claims.count() * 100) if last_year_claims.count() > 0 else 0,
            }
        }
        
        # === MONTHLY BREAKDOWN - SIMPLIFIED ===
        monthly_breakdown = []
        for month in range(1, 13):
            month_claims = shipments.filter(
                Intend_Claim_Date__year=current_year,
                Intend_Claim_Date__month=month
            )
            monthly_breakdown.append({
                'month': calendar.month_name[month],
                'count': month_claims.count(),
                'value': float(month_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                'savings': float(month_claims.aggregate(total=models.Sum('Total_Savings'))['total'] or 0),
                'settled_count': month_claims.filter(Settlement_Status='SETTLED').count(),
            })
        
        # === STATUS ANALYSIS - SIMPLIFIED ===
        status_data = []
        status_choices = ['OPEN', 'PENDING', 'CLOSED']
        for status in status_choices:
            status_claims = shipments.filter(Status=status)
            status_data.append({
                'status': status,
                'count': status_claims.count(),
                'total_value': float(status_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                'avg_value': float(status_claims.aggregate(avg=models.Avg('Claimed_Amount'))['avg'] or 0),
                'percentage': (status_claims.count() / total_claims * 100) if total_claims > 0 else 0
            })
        
        # === SETTLEMENT ANALYSIS ===
        settlement_data = []
        settlement_choices = ['SETTLED', 'NOT_SETTLED', 'PARTIAL']
        for settlement in settlement_choices:
            settlement_claims = shipments.filter(Settlement_Status=settlement)
            settlement_data.append({
                'settlement_status': settlement,
                'count': settlement_claims.count(),
                'total_value': float(settlement_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                'avg_value': float(settlement_claims.aggregate(avg=models.Avg('Claimed_Amount'))['avg'] or 0),
                'percentage': (settlement_claims.count() / total_claims * 100) if total_claims > 0 else 0
            })
        
        # === BRANCH ANALYSIS - SIMPLIFIED ===
        branches = shipments.values_list('Branch', flat=True).distinct()
        branch_data = []
        for branch in branches:
            if branch:  # Skip None values
                branch_claims = shipments.filter(Branch=branch)
                branch_data.append({
                    'branch': branch,
                    'count': branch_claims.count(),
                    'total_value': float(branch_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                    'avg_value': float(branch_claims.aggregate(avg=models.Avg('Claimed_Amount'))['avg'] or 0),
                    'settled_count': branch_claims.filter(Settlement_Status='SETTLED').count(),
                    'settlement_rate': (branch_claims.filter(Settlement_Status='SETTLED').count() / branch_claims.count() * 100) if branch_claims.count() > 0 else 0
                })
        branch_data = sorted(branch_data, key=lambda x: x['total_value'], reverse=True)
        
        # === CLIENT ANALYSIS - SIMPLIFIED ===
        client_data = []
        # Only get clients that actually have shipments
        clients_with_shipments = shipments.values_list('client_id', flat=True).distinct()
        for client_id in clients_with_shipments:
            if client_id:  # Skip None values
                try:
                    client = Client.objects.get(id=client_id)
                    client_claims = shipments.filter(client_id=client_id)
                    client_data.append({
                        'client_name': client.name,
                        'client_id': client.client_id,
                        'count': client_claims.count(),
                        'total_value': float(client_claims.aggregate(total=models.Sum('Claimed_Amount'))['total'] or 0),
                        'avg_value': float(client_claims.aggregate(avg=models.Avg('Claimed_Amount'))['avg'] or 0),
                        'settled_count': client_claims.filter(Settlement_Status='SETTLED').count(),
                        'settlement_rate': (client_claims.filter(Settlement_Status='SETTLED').count() / client_claims.count() * 100) if client_claims.count() > 0 else 0
                    })
                except Client.DoesNotExist:
                    continue
        client_data = sorted(client_data, key=lambda x: x['total_value'], reverse=True)[:20]  # Top 20 clients
        
        # === PAYMENT ANALYSIS ===
        payment_analysis = {
            'iscm_awa': {
                'total': total_paid_iscm,
                'count': shipments.filter(Amount_Paid_By_Awa__gt=0).count(),
                'percentage': (total_paid_iscm / total_paid_all * 100) if total_paid_all > 0 else 0,
            },
            'carrier': {
                'total': total_paid_carrier,
                'count': shipments.filter(Amount_Paid_By_Carrier__gt=0).count(),
                'percentage': (total_paid_carrier / total_paid_all * 100) if total_paid_all > 0 else 0,
            },
            'insurance': {
                'total': total_paid_insurance,
                'count': shipments.filter(Amount_Paid_By_Insurance__gt=0).count(),
                'percentage': (total_paid_insurance / total_paid_all * 100) if total_paid_all > 0 else 0,
            }
        }
        
        # === EFFICIENCY METRICS ===
        intent_claims = shipments.filter(Intent_To_Claim='YES').count()
        formal_claims = shipments.filter(Formal_Claim_Received='YES').count()
        settled_claims = shipments.filter(Settlement_Status='SETTLED').count()
        
        efficiency_metrics = {
            'intent_to_formal_rate': (formal_claims / intent_claims * 100) if intent_claims > 0 else 0,
            'overall_success_rate': (settled_claims / total_claims * 100) if total_claims > 0 else 0,
            'claims_requiring_formal': formal_claims,
            'total_intent_claims': intent_claims,
            'total_settled_claims': settled_claims,
        }
        
        context = {
            # Basic metrics
            'total_claims': total_claims,
            'total_value': total_value,
            'total_paid_iscm': total_paid_iscm,
            'total_paid_carrier': total_paid_carrier,
            'total_paid_insurance': total_paid_insurance,
            'total_paid_all': total_paid_all,
            'total_savings': total_savings,
            'total_exposure': total_exposure,
            'avg_claim_value': avg_claim_value,
            'avg_savings_per_claim': avg_savings_per_claim,
            'avg_exposure_per_claim': avg_exposure_per_claim,
            
            # Rates and percentages
            'savings_rate': round(savings_rate, 2),
            'recovery_rate': round(recovery_rate, 2),
            'exposure_rate': round(exposure_rate, 2),
            
            # Analysis data
            'time_analysis': time_analysis,
            'yearly_data': yearly_data,
            'monthly_breakdown': monthly_breakdown,
            'status_data': status_data,
            'settlement_data': settlement_data,
            'branch_data': branch_data,
            'client_data': client_data,
            'payment_analysis': payment_analysis,
            'efficiency_metrics': efficiency_metrics,
        }
        
        return render(request, 'main/analytics_dashboard.html', context)
    
    except Exception as e:
        # Log the error and show a friendly message
        print(f"Analytics dashboard error: {e}")
        messages.error(request, f"Error loading analytics dashboard: {str(e)}")
        return redirect('home')


# Remove the old _calculate_avg_processing_time function since it's causing issues

def _calculate_avg_processing_time(queryset):
    """Helper function to calculate average processing time."""
    processing_times = []
    for shipment in queryset.filter(
        Intend_Claim_Date__isnull=False,
        Formal_Claim_Date_Received__isnull=False
    ):
        if shipment.Formal_Claim_Date_Received > shipment.Intend_Claim_Date:
            days = (shipment.Formal_Claim_Date_Received - shipment.Intend_Claim_Date).days
            processing_times.append(days)
    
    return sum(processing_times) / len(processing_times) if processing_times else 0