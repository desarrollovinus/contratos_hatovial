<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <!--[if lt IE 9]>
        <script src=”http://HTML5shim.googlecode.com/svn/trunk/HTML5.js”>
        </script>
        <![endif]-->
        <?php $this->load->view('includes/header'); ?>
    </head>
    <body>
        <div class="container_12">
            <section id="cuerpo">
                <article id="contenido">
                    <?php $this->load->view($contenido_principal); ?>
                </article>
            </section>
            <?php $this->load->view('includes/footer'); ?>
        </div>
    </body>
</html>