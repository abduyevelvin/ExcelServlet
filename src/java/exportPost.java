import java.io.IOException;
import java.io.PrintWriter;
import java.util.Date;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public class exportPost extends HttpServlet {

    String dbURL;

    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        response.setContentType("text/html;charset=UTF-8");

        LoadDataExcelFromOracl loadexcelfile = new LoadDataExcelFromOracl();

        PrintWriter out = response.getWriter();

        String fileName = "export" + ((new Date()).getTime()) + ".xlsx";
        String dbDriver = "oracle.jdbc.driver.OracleDriver";
        String SQLQuery = request.getParameter("queryP");
        String TypeCon = request.getParameter("ConType");

        switch (TypeCon) {
            case "1":
                dbURL = "jdbc:oracle:thin: (1)";
                break;
            case "2":
                dbURL = "jdbc:oracle:thin: (2)";
                break;
            case "3":
                dbURL = "jdbc:oracle:thin:(3)";
                break;
        }

        fileName = loadexcelfile.LoadDataExcel(SQLQuery,
                "borc",
                "user",
                "pass",
                dbDriver,
                dbURL,
                "D:\\",
                fileName
        );

        out.println("<html>");
        out.println("<head>");
        out.println("<title>Excel</title>");
        out.println("</head>");
        out.println("<body bgcolor=\"white\">");
        out.println("<a href=(download URL from folder)" + fileName + "><h3>Download file</h3></a>");
        out.println("</body>");
        out.println("</html>");
    }

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /**
     * Handles the HTTP <code>GET</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Handles the HTTP <code>POST</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>
}