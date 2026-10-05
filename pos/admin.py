from django.contrib import admin

from .models import Customer, Payment, Product, Sale, SaleItem


class SaleItemInline(admin.TabularInline):
    model = SaleItem
    extra = 0
    readonly_fields = ("product", "product_name", "unit_price", "quantity", "line_total")


@admin.register(Sale)
class SaleAdmin(admin.ModelAdmin):
    list_display = ("receipt_number", "created_at", "customer", "total", "status")
    list_filter = ("status", "created_at")
    search_fields = ("receipt_number", "customer__name")
    inlines = (SaleItemInline,)
    readonly_fields = ("receipt_number", "created_at", "subtotal", "total")


@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ("name", "sku", "price", "stock", "is_active")
    search_fields = ("name", "sku")
    list_filter = ("is_active",)


@admin.register(Customer)
class CustomerAdmin(admin.ModelAdmin):
    list_display = ("name", "email", "phone", "created_at")
    search_fields = ("name", "email", "phone")


@admin.register(Payment)
class PaymentAdmin(admin.ModelAdmin):
    list_display = ("sale", "method", "amount", "paid_at", "transaction_reference")
    list_filter = ("method", "paid_at")
