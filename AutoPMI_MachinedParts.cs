
using System;
using System.Collections.Generic;
using NXOpen;
using NXOpen.Annotations;
using NXOpen.Features;
using NXOpen.UF;

public class AutoPMI_MachinedParts
{
    private static Session session = Session.GetSession();
    private static UFSession uf = UFSession.GetUFSession();

    public static void Main(string[] args)
    {
        ListingWindow lw = session.ListingWindow;
        lw.Open();

        Part part = session.Parts.Work;

        if (part == null)
        {
            lw.WriteLine("ERROR: Open a machined part first.");
            return;
        }

        // Assumptions:
        // 1. Hole axes are parallel to absolute Z.
        // 2. WCS XY is aligned with the part's top view.
        // 3. Holes are modeled as NX hole features.
        // 4. PMI/GD&T licensing is available.

        int found = 0;
        int created = 0;
        int existing = 0;
        int failed = 0;

        HashSet<Tag> processedFaces = new HashSet<Tag>();

        Session.UndoMarkId mark =
            session.SetUndoMark(
                Session.MarkVisibility.Visible,
                "Automatic Machined-Part PMI");

        lw.WriteLine("====================================");
        lw.WriteLine("AUTOMATIC PMI - MACHINED COMPONENTS");
        lw.WriteLine("Part: " + part.Leaf);
        lw.WriteLine("====================================");

        foreach (Feature feature in part.Features)
        {
            string type = feature.FeatureType.ToUpperInvariant();

            // Recognize standard NX hole feature types.
            // Expand this filter for custom modeling practices.
            if (!type.Contains("HOLE"))
                continue;

            Face[] faces;

            try
            {
                faces = feature.GetFaces();
            }
            catch (Exception ex)
            {
                lw.WriteLine(
                    "Cannot read feature " + feature.Name +
                    ": " + ex.Message);
                continue;
            }

            foreach (Face face in faces)
            {
                if (face.SolidFaceType !=
                    Face.FaceType.Cylindrical)
                    continue;

                if (!processedFaces.Add(face.Tag))
                    continue;

                // Read cylinder axis, radius and location.
                int faceType;
                int normalDirection;
                double radius;
                double radiusData;

                double[] axisPoint = new double[3];
                double[] axisDirection = new double[3];
                double[] boundingBox = new double[6];

                try
                {
                    uf.Modl.AskFaceData(
                        face.Tag,
                        out faceType,
                        axisPoint,
                        axisDirection,
                        boundingBox,
                        out radius,
                        out radiusData,
                        out normalDirection);
                }
                catch (Exception ex)
                {
                    lw.WriteLine(
                        "Cannot measure face: " + ex.Message);
                    failed++;
                    continue;
                }

                // This pilot handles Z-axis holes only.
                if (Math.Abs(axisDirection[2]) < 0.98)
                {
                    lw.WriteLine(
                        "Skipped non-Z-axis cylindrical face.");
                    continue;
                }

                if (radius <= 0.0)
                    continue;

                // Find a complete circular edge of this face.
                Edge circle = null;

                foreach (Edge edge in face.GetEdges())
                {
                    if (edge.SolidEdgeType !=
                        Edge.EdgeType.Circular)
                        continue;

                    Point3d start;
                    Point3d end;

                    edge.GetVertices(out start, out end);

                    double separation = Math.Sqrt(
                        Math.Pow(start.X - end.X, 2) +
                        Math.Pow(start.Y - end.Y, 2) +
                        Math.Pow(start.Z - end.Z, 2));

                    // A closed circular edge has coincident
                    // start and end vertices.
                    if (separation < 0.0001)
                    {
                        circle = edge;
                        break;
                    }
                }

                if (circle == null)
                {
                    lw.WriteLine(
                        "Skipped face: no complete circular edge.");
                    continue;
                }

                found++;

                if (HasDiameterPmi(part, face))
                {
                    existing++;
                    lw.WriteLine(
                        "Existing diameter PMI found on face.");
                    continue;
                }

                try
                {
                    CreateDiameterPmi(
                        part,
                        circle,
                        axisPoint,
                        radius);

                    created++;

                    lw.WriteLine(
                        "Created PMI diameter: " +
                        (2.0 * radius).ToString("F3") +
                        " part units");
                }
                catch (Exception ex)
                {
                    failed++;

                    lw.WriteLine(
                        "PMI creation failed for " +
                        feature.Name + ": " + ex.Message);
                }
            }
        }

        try
        {
            session.UpdateManager.DoUpdate(mark);
        }
        catch (Exception ex)
        {
            lw.WriteLine(
                "Update warning: " + ex.Message);
        }

        lw.WriteLine("");
        lw.WriteLine("=========== SUMMARY ===========");
        lw.WriteLine("Hole faces detected: " + found);
        lw.WriteLine("PMI created: " + created);
        lw.WriteLine("Already dimensioned: " + existing);
        lw.WriteLine("Failures: " + failed);
        lw.WriteLine("================================");

        lw.WriteLine(
            "Review all dimensions before releasing MBD.");
        lw.WriteLine(
            "The journal does not assign GD&T or tolerances.");
    }

    private static bool HasDiameterPmi(
        Part part,
        Face cylindricalFace)
    {
        HashSet<Tag> faceEdges = new HashSet<Tag>();

        foreach (Edge edge in cylindricalFace.GetEdges())
            faceEdges.Add(edge.Tag);

        foreach (Dimension dim in part.Dimensions)
        {
            if (!(dim is PmiDiameterDimension))
                continue;

            for (int index = 1;
                 index <= dim.NumberOfAssociativities;
                 index++)
            {
                try
                {
                    Associativity assoc =
                        dim.GetAssociativity(index);

                    if (assoc == null)
                        continue;

                    if (assoc.FirstObject != null &&
                        faceEdges.Contains(
                            assoc.FirstObject.Tag))
                        return true;

                    if (assoc.SecondObject != null &&
                        faceEdges.Contains(
                            assoc.SecondObject.Tag))
                        return true;
                }
                catch
                {
                    // Ignore unreadable associations.
                }
            }
        }

        return false;
    }

    private static void CreateDiameterPmi(
        Part part,
        Edge circularEdge,
        double[] center,
        double radius)
    {
        using (DimensionData dimensionData =
            part.Annotations.NewDimensionData())
        using (PmiData pmiData =
            part.Annotations.NewPmiData())
        using (Associativity associativity =
            part.Annotations.NewAssociativity())
        {
            associativity.FirstObject = circularEdge;
            associativity.SecondObject = null;
            associativity.ObjectView =
                part.Views.WorkView;

            dimensionData.SetAssociativity(
                1,
                new Associativity[] { associativity });

            // Place the PMI on the WCS XY plane.
            Xform plane =
                part.Annotations.GetDefaultAnnotationPlane(
                    PmiDefaultPlane.XyOfWcs);

            // Initial text placement. The engineer may
            // reposition it to match company standards.
            double offset = Math.Max(
                radius * 1.5, 5.0);

            Point3d origin = new Point3d(
                center[0] + offset,
                center[1] + offset,
                center[2]);

            PmiDiameterDimension dimension =
                part.Dimensions.CreatePmiDiameterDimension(
                    dimensionData,
                    pmiData,
                    plane,
                    origin);

            if (dimension == null)
                throw new Exception(
                    "NX returned no PMI dimension.");
        }
    }

    public static int GetUnloadOption(string dummy)
    {
        return (int)
            Session.LibraryUnloadOption.Immediately;
    }
}
