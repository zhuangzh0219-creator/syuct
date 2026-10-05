import json

from django.test import TestCase
from django.urls import reverse

from .models import Customer, Payment, Product, Sale


class PosWorkflowTests(TestCase):
    def setUp(self):
        self.product = Product.objects.create(name="テスト商品", sku="T-001", price=500, stock=4)
        self.customer = Customer.objects.create(name="テスト顧客")

    def test_checkout_uses_database_price_and_updates_stock(self):
        response = self.client.post(
            reverse("checkout"),
            {
                "cart": json.dumps([{"id": self.product.pk, "quantity": 2, "price": 1}]),
                "customer_id": str(self.customer.pk),
                "payment_method": Payment.Method.PAYPAY,
            },
        )

        self.assertEqual(response.status_code, 200)
        data = response.json()
        sale = Sale.objects.get()
        self.product.refresh_from_db()
        self.assertEqual(sale.total, 1000)
        self.assertEqual(sale.customer, self.customer)
        self.assertEqual(sale.items.get().line_total, 1000)
        self.assertEqual(sale.payment.method, Payment.Method.PAYPAY)
        self.assertEqual(self.product.stock, 2)
        self.assertIn(f"/receipts/{sale.pk}/", data["receipt_url"])

    def test_checkout_rejects_insufficient_stock_without_saving_sale(self):
        response = self.client.post(
            reverse("checkout"),
            {
                "cart": json.dumps([{"id": self.product.pk, "quantity": 5}]),
                "customer_id": "",
                "payment_method": Payment.Method.CASH,
            },
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(Sale.objects.count(), 0)
        self.product.refresh_from_db()
        self.assertEqual(self.product.stock, 4)

    def test_receipt_and_report_csv_are_available(self):
        response = self.client.post(
            reverse("checkout"),
            {
                "cart": json.dumps([{"id": self.product.pk, "quantity": 1}]),
                "customer_id": "",
                "payment_method": Payment.Method.ALIPAY,
            },
        )
        sale = Sale.objects.get()

        receipt = self.client.get(reverse("receipt_download", args=[sale.pk]))
        report = self.client.get(reverse("report_export"))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(receipt.status_code, 200)
        self.assertIn("領収書番号".encode("utf-8"), receipt.content)
        self.assertEqual(report.status_code, 200)
        self.assertIn(sale.receipt_number.encode("utf-8"), report.content)
        self.assertContains(self.client.get(reverse("reports")), "支付宝 / Alipay")
