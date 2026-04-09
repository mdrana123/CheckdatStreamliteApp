import os
import random
import shutil
 
# ===================== CONFIG =====================
DATASET_DIR = r"C:\Users\alamr\OneDrive\Desktop\checkdat_2\Checkdat_Stampel_AI\DataSet"
IMAGE_DIR = os.path.join(DATASET_DIR, "images")
LABEL_DIR = os.path.join(DATASET_DIR, "labels")
 
TRAIN_RATIO = 0.8           # 80% train, 20% val
SEED = 42                   # random seed for reproducibility
 
# ==================================================
 
random.seed(SEED)
 
images = [f for f in os.listdir(IMAGE_DIR) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
random.shuffle(images)
 
train_size = int(len(images) * TRAIN_RATIO)
train_files = images[:train_size]
val_files = images[train_size:]
 
def move_pairs(file_list, dest):
    os.makedirs(os.path.join(IMAGE_DIR, dest), exist_ok=True)
    os.makedirs(os.path.join(LABEL_DIR, dest), exist_ok=True)
 
    for img in file_list:
        base = os.path.splitext(img)[0]
        lbl = base + ".txt"
 
        shutil.copy(os.path.join(IMAGE_DIR, img), os.path.join(IMAGE_DIR, dest, img))
        shutil.copy(os.path.join(LABEL_DIR, lbl), os.path.join(LABEL_DIR, dest, lbl))
 
move_pairs(train_files, "train")
move_pairs(val_files, "val")
 
print(f"✅ TRAIN: {len(train_files)} images")
print(f"✅ VAL  : {len(val_files)} images")
print("\n🎉 Automatic split completed.")
 
 