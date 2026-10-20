import unittest


class PruebasSistema(unittest.TestCase):

    def test_calcular_total(self):
        precio = 10
        cantidad = 3

        total = precio * cantidad

        self.assertEqual(total, 30)

    def test_existencia_producto(self):
        existencia = 20
        cantidad_vendida = 3

        existencia_final = existencia - cantidad_vendida

        self.assertEqual(existencia_final, 17)

    def test_devolver_existencia(self):
        existencia = 17
        cantidad_devuelta = 3

        existencia_final = existencia + cantidad_devuelta

        self.assertEqual(existencia_final, 20)

    def test_venta_total(self):
        ventas = [30, 20, 15]

        total = sum(ventas)

        self.assertEqual(total, 65)


if __name__ == "__main__":
    unittest.main()