from ultralytics import YOLO

model = YOLO("yolov8m.pt")

# Train (CPU)
model.train(
    data="data.yaml",
    epochs=80,
    imgsz=1248,
    batch=2,          # 🔑 CPU-safe for 2048
    device="cpu",
    patience=30,
    workers=0         # avoids Windows multiprocessing issues
)

# Validate using same imgsz
metrics = model.val(imgsz=1248, device="cpu")
print(metrics)

# Inference (start low conf to verify boxes)
results = model.predict(
    source="DataSet/images/val",
    imgsz=1248,
    conf=0.05,        # 🔑 start low; raise later (0.25) after it works
    device="cpu",
    verbose=False
)
