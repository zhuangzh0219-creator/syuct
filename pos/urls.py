from django.urls import path

from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("pos/", views.pos_screen, name="pos"),
    path("checkout/", views.checkout, name="checkout"),
    path("products/", views.products, name="products"),
    path("customers/", views.customers, name="customers"),
    path("receipts/", views.receipts, name="receipts"),
    path("receipts/<int:pk>/", views.receipt_detail, name="receipt_detail"),
    path("receipts/<int:pk>/download/", views.receipt_download, name="receipt_download"),
    path("reports/", views.reports, name="reports"),
    path("reports/export/", views.report_export, name="report_export"),
]
