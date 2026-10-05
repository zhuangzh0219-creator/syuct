import csv
import json
from datetime import date, datetime, time, timedelta

from django.contrib import messages
from django.db import transaction
from django.db.models import Count, Sum
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.csrf import ensure_csrf_cookie
from django.views.decorators.http import require_POST

from .models import Customer, Payment, Product, Sale, SaleItem, create_demo_reference


def dashboard(request):
    today = timezone.localdate()
    today_sales = Sale.objects.filter(created_at__date=today, status=Sale.Status.PAID)
    return render(
        request,
        "pos/dashboard.html",
        {
            "today_total": today_sales.aggregate(total=Sum("total"))["total"] or 0,
            "today_count": today_sales.count(),
            "customer_count": Customer.objects.count(),
            "low_stock_count": Product.objects.filter(is_active=True, stock__lte=5).count(),
            "recent_sales": Sale.objects.filter(status=Sale.Status.PAID).select_related("customer", "payment")[:6],
        },
    )


@ensure_csrf_cookie
def pos_screen(request):
    return render(
        request,
        "pos/pos.html",
        {
            "products": Product.objects.filter(is_active=True, stock__gt=0),
            "customers": Customer.objects.all(),
            "payment_methods": Payment.Method.choices,
        },
    )


@require_POST
def checkout(request):
    try:
        cart = json.loads(request.POST.get("cart", ""))
    except (json.JSONDecodeError, TypeError):
        return JsonResponse({"error": "カートの内容を読み取れませんでした。"}, status=400)

    if not isinstance(cart, list) or not cart:
        return JsonResponse({"error": "商品をカートに追加してください。"}, status=400)

    method = request.POST.get("payment_method", "")
    if method not in Payment.Method.values:
        return JsonResponse({"error": "支払方法を選択してください。"}, status=400)

    customer_id = request.POST.get("customer_id", "").strip()
    customer = None
    if customer_id:
        try:
            customer = Customer.objects.get(pk=int(customer_id))
        except (ValueError, Customer.DoesNotExist):
            return JsonResponse({"error": "選択した顧客が見つかりません。"}, status=400)

    quantities = {}
    for row in cart:
        if not isinstance(row, dict):
            return JsonResponse({"error": "カートの商品情報が不正です。"}, status=400)
        try:
            product_id = int(row["id"])
            quantity = int(row["quantity"])
        except (KeyError, TypeError, ValueError):
            return JsonResponse({"error": "カートの商品情報が不正です。"}, status=400)
        if quantity < 1:
            return JsonResponse({"error": "数量は1以上にしてください。"}, status=400)
        quantities[product_id] = quantities.get(product_id, 0) + quantity

    with transaction.atomic():
        products = {
            product.pk: product
            for product in Product.objects.select_for_update().filter(pk__in=quantities, is_active=True)
        }
        if set(products) != set(quantities):
            return JsonResponse({"error": "販売できない商品が含まれています。画面を更新してください。"}, status=400)
        for product_id, quantity in quantities.items():
            if products[product_id].stock < quantity:
                return JsonResponse({"error": f"「{products[product_id].name}」の在庫が不足しています。"}, status=400)

        amount = sum(products[product_id].price * quantity for product_id, quantity in quantities.items())
        sale = Sale.objects.create(subtotal=amount, total=amount, customer=customer)
        for product_id, quantity in quantities.items():
            product = products[product_id]
            SaleItem.objects.create(
                sale=sale,
                product=product,
                product_name=product.name,
                unit_price=product.price,
                quantity=quantity,
                line_total=product.price * quantity,
            )
            product.stock -= quantity
            product.save(update_fields=["stock"])
        sale.receipt_number = f"R-{timezone.localdate():%Y%m%d}-{sale.pk:06d}"
        sale.save(update_fields=["receipt_number"])
        Payment.objects.create(
            sale=sale,
            method=method,
            amount=amount,
            transaction_reference=create_demo_reference(),
        )

    return JsonResponse(
        {
            "ok": True,
            "message": "デモ決済が完了しました。実際の決済は行われていません。",
            "receipt_url": f"/receipts/{sale.pk}/",
        }
    )


def products(request):
    if request.method == "POST":
        name = request.POST.get("name", "").strip()
        sku = request.POST.get("sku", "").strip()
        try:
            price = int(request.POST.get("price", ""))
            stock = int(request.POST.get("stock", ""))
        except ValueError:
            messages.error(request, "価格と在庫数には整数を入力してください。")
        else:
            if not name or not sku or price < 0 or stock < 0:
                messages.error(request, "商品名・商品コードを入力し、価格と在庫数は0以上にしてください。")
            elif Product.objects.filter(sku=sku).exists():
                messages.error(request, "この商品コードはすでに登録されています。")
            else:
                Product.objects.create(name=name, sku=sku, price=price, stock=stock)
                messages.success(request, "商品を登録しました。")
                return redirect("products")
    return render(request, "pos/products.html", {"products": Product.objects.all()})


def customers(request):
    if request.method == "POST":
        name = request.POST.get("name", "").strip()
        email = request.POST.get("email", "").strip()
        phone = request.POST.get("phone", "").strip()
        if not name:
            messages.error(request, "顧客名を入力してください。")
        elif len(name) > 120 or len(email) > 254 or len(phone) > 30:
            messages.error(request, "入力内容が長すぎます。")
        else:
            Customer.objects.create(name=name, email=email, phone=phone)
            messages.success(request, "顧客を登録しました。")
            return redirect("customers")
    return render(request, "pos/customers.html", {"customers": Customer.objects.all()})


def receipts(request):
    sales = Sale.objects.filter(status=Sale.Status.PAID).select_related("customer", "payment")
    return render(request, "pos/receipts.html", {"sales": sales})


def receipt_detail(request, pk):
    sale = get_object_or_404(
        Sale.objects.filter(status=Sale.Status.PAID).select_related("customer", "payment").prefetch_related("items"),
        pk=pk,
    )
    return render(request, "pos/receipt_detail.html", {"sale": sale})


def receipt_download(request, pk):
    sale = get_object_or_404(
        Sale.objects.filter(status=Sale.Status.PAID).select_related("customer", "payment").prefetch_related("items"),
        pk=pk,
    )
    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = f'attachment; filename="{sale.receipt_number}.csv"'
    response.write("\ufeff")
    writer = csv.writer(response)
    writer.writerow(["領収書番号", sale.receipt_number])
    writer.writerow(["販売日時", timezone.localtime(sale.created_at).strftime("%Y-%m-%d %H:%M")])
    writer.writerow(["顧客", sale.customer.name if sale.customer else "ゲスト"])
    writer.writerow(["商品", "単価（税込円）", "数量", "金額（税込円）"])
    for item in sale.items.all():
        writer.writerow([item.product_name, item.unit_price, item.quantity, item.line_total])
    writer.writerow(["支払方法", sale.payment.get_method_display()])
    writer.writerow(["合計（税込円）", sale.total])
    return response


def _report_range(request):
    today = timezone.localdate()
    try:
        start = date.fromisoformat(request.GET.get("start", ""))
    except ValueError:
        start = today.replace(day=1)
    try:
        end = date.fromisoformat(request.GET.get("end", ""))
    except ValueError:
        end = today
    if start > end:
        start, end = end, start
    start_dt = timezone.make_aware(datetime.combine(start, time.min))
    end_dt = timezone.make_aware(datetime.combine(end + timedelta(days=1), time.min))
    sales = Sale.objects.filter(status=Sale.Status.PAID, created_at__gte=start_dt, created_at__lt=end_dt)
    return start, end, sales


def reports(request):
    start, end, sales = _report_range(request)
    payment_totals = list(
        Payment.objects.filter(sale__in=sales)
        .values("method")
        .annotate(count=Count("id"), total=Sum("amount"))
        .order_by("method")
    )
    method_labels = dict(Payment.Method.choices)
    for row in payment_totals:
        row["method_label"] = method_labels[row["method"]]
    top_products = (
        SaleItem.objects.filter(sale__in=sales)
        .values("product_name")
        .annotate(quantity=Sum("quantity"), total=Sum("line_total"))
        .order_by("-quantity", "product_name")[:8]
    )
    return render(
        request,
        "pos/reports.html",
        {
            "start": start,
            "end": end,
            "sales_count": sales.count(),
            "sales_total": sales.aggregate(total=Sum("total"))["total"] or 0,
            "payment_totals": payment_totals,
            "top_products": top_products,
        },
    )


def report_export(request):
    start, end, sales = _report_range(request)
    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = f'attachment; filename="sales-report-{start}-{end}.csv"'
    response.write("\ufeff")
    writer = csv.writer(response)
    writer.writerow(["領収書番号", "販売日時", "顧客", "支払方法", "合計（税込円）"])
    for sale in sales.select_related("customer", "payment"):
        writer.writerow(
            [
                sale.receipt_number,
                timezone.localtime(sale.created_at).strftime("%Y-%m-%d %H:%M"),
                sale.customer.name if sale.customer else "ゲスト",
                sale.payment.get_method_display(),
                sale.total,
            ]
        )
    return response
