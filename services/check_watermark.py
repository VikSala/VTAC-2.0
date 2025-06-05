"""from skimage.metrics import structural_similarity as ssim
import cv2

# Cargar imágenes
img1 = cv2.imread('sin_marca.jpg')
img2 = cv2.imread('con_marca.jpg')

# Convertir a escala de grises
gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)

# Calcular SSIM
score, diff = ssim(gray1, gray2, full=True)
print(f"SSIM: {score}")
if score < 0.9: print("No hay marca de agua entre los dos archivos...")
else: print("Si hay marca de agua entre los dos archivos!!!")"""

