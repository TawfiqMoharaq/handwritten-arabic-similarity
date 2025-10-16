# Search for Similarity in Handwritten Arabic Letters 🖋️

This repository contains my official solution for **KAUST Academy — Question 2**.  
The task focuses on building an **image search system** to find the top 5 most similar Arabic handwritten characters using **EfficientNet-B3** and **cosine similarity**.

---

## 🧩 Task Summary
> Given a handwritten Arabic character (e.g., "ب"), find the 5 most similar images from a dataset of handwritten Arabic characters.

### Steps:
1. Load query image and dataset images.
2. Build a custom PyTorch dataset.
3. Extract image features using a pretrained **EfficientNet-B3** model.
4. Reduce dimensionality using **PCA**.
5. Compute **cosine similarity** between query and dataset features.
6. Display the top-5 most similar images with their similarity scores.

---

## 🛠️ Technologies Used
- Python
- PyTorch & Torchvision
- Scikit-learn
- NumPy, Matplotlib
- EfficientNet-B3 (pretrained)
- KaggleHub Dataset Loader

---

## 📊 Dataset
- [mohammad2012191/arabic-chars](https://www.kaggle.com/datasets/mohammad2012191/arabic-chars)
- [mohammad2012191/character](https://www.kaggle.com/datasets/mohammad2012191/character)

---

## 🧠 Results
Displays:
- Query image.
- Top-5 most similar handwritten Arabic characters with cosine similarity percentages.


