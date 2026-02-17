# Aerospace Section Analysis Tool & Physics Engine (VBA)

This repository contains a specialized engineering tool developed to perform **Computational Geometry** and **Physical Property Analysis** (Center of Mass, Moment of Inertia) for 2D cross-sections of 3D objects.

It functions as a lightweight CAD/Physics engine running entirely within Microsoft Excel, built with **VBA (Visual Basic for Applications)**.

---

## 📖 The Story: "The 13-Day Sprint"

This project is the result of a rigorous engineering challenge.

* **Timeline:** Written in exactly **13 days**.
* **Starting Point:** I started with **zero knowledge of VBA**. I learned the language syntax and capabilities simultaneously while developing the core logic.
* **The Constraint:** **NO Artificial Intelligence (AI) assistance was used.** Every algorithm, every line of logic, and every bug fix was handled manually.
* **The Process:** The complex geometric algorithms (such as Ray Casting for point-in-polygon detection and edge intersection logic) were not copied from libraries. They were mathematically derived and sketched by hand on pages of a physical notebook.

> **Note on the "Lost Notebook":** The physical notebook containing pages of hand-drawn edge-case scenarios (e.g., *"what if a point lies exactly on the edge?"*, *"tangent shapes"*) is currently lost. If found, scans of these derivations will be uploaded to the `Media/` folder to visualize the mathematical thought process behind the code.

### ⚠️ State of the Code: "Raw & Unfiltered"
This repository represents the project **exactly as it was found** after the deadline.
* **No Refactoring:** To preserve the authenticity of the 13-day sprint, the code has not been polished or modernized.
* **Complexity:** The code structure reflects a race against time. You may find complex nested loops or unconventional solutions—these are artifacts of solving high-level math problems under extreme time pressure.
* **Language:** Variable and function names are largely in **Turkish** (e.g., `KesisimBul` instead of `FindIntersection`) as I coded in my native language to maximize development speed.

---

## 🚀 Key Technical Features

Despite the constraints, the tool successfully implements advanced engineering concepts:

### 1. Computational Geometry Engine
* **Ray Casting Algorithm:** Implements a custom "Point in Polygon" algorithm to determine if a specific point lies within a complex, irregular shape.
* **Edge Intersection Logic:** Calculates exact intersection points between slicing planes and 3D geometries to generate 2D cross-sections.
* **Collision Detection:** Identifies overlaps between independent parts (Vertex-based detection).

### 2. Physics & Inertia Calculation
* **Center of Mass:** Calculates the centroid of composite shapes.
* **Moment of Inertia ($I_{xx}, I_{yy}, I_{zz}$):** Uses the **Parallel Axis Theorem (Steiner's Theorem)** to calculate inertia relative to the center of gravity.
* **Product of Inertia ($I_{xy}, I_{xz}, I_{yz}$):** Calculates products of inertia to assess symmetry and rotational stability.

### 3. Object-Oriented Design (OOP) in VBA
* Uses a custom `Nokta` (Point) class to handle 3D coordinates ($x, y, z$) as objects rather than simple arrays, enabling cleaner data manipulation.

---

## 📂 Project Structure

The project logic is modularized for version control:

* **`App_Main.bas`**: The entry point and main execution loop. Handles the workflow logic.
* **`Core_Geometry.bas`**: The mathematical heart of the project. Contains `Ray Casting`, `Intersection`, and `Shape Cutting` algorithms.
* **`Manager_Spatial.bas`**: Handles coordinate systems, axis definitions, and piece identification.
* **`Validation.bas`**: Regex-based validation for user input functions.
* **`IO_Handler.bas`**: Manages Excel read/write operations for reporting results.
* **`Tests.bas`**: Unit tests for edge cases (e.g., points on edges, touching shapes).
* **`Nokta.cls`**: Class module defining the 3D Point object.


---

## 📝 Known Issues & Limitations

* **Inertia Approximation:** Currently, the inertia calculation treats discretized pieces as **Point Masses** ($I = md^2$). The local inertia of the piece itself ($I_{local}$) is negligible for small mesh sizes but technically omitted.
* **Collision Precision:** Collision detection is primarily vertex-based. Edge-to-edge intersection without vertex penetration might be missed in rare scenarios.
* **Variable Naming:** As mentioned, internal variables are in Turkish (`yuz`, `parca`, `hesapla`).

---

## 👨‍💻 Author

**Yusuf Furkan Umutlu**
*Computer Engineering Graduate | Embedded Systems Enthusiast*

*This project stands as a testament to rapid learning, algorithmic thinking, and engineering problem-solving under strict constraints.*