#Requires AutoHotkey >=v2
#SingleInstance Force
#Include <OCR>  ; https://github.com/Descolada/OCR

^Esc::ExitApp

ocr_from_rect(x, y, w, h)
{
  return OCR.FromRect(x, y, w, h, "en-DE", 6)
}

try_open_login_page(login_page_name)
{
  if (not WinActive(login_page_name))
    Run("explorer https://webmail.hs-merseburg.de/src/login.php")
    ; "explorer" runs the system default browser
}

do_login(login_page_name)
{
  WinWaitActive(login_page_name)

  Click("919 490 Left Down")
  Click("919 490 Left Up")
  SendText("username")
  SendInput("{Tab}")  ; navigate to password input
  SendText("password")
  SendInput("{Enter}")
}

do_search(web_page_name)
{
  Loop
  {
    found_x := 0
    found_y := 0

    KeyWait("F3", "D") ; wait for keypress down

    WinActivate(web_page_name)

    If (A_Index = 1)
    {
      do_search_sequence(found_x, found_y, "yes")
      do_search_and_copy_delivery_date(found_x, found_y)
    }
    else
    {
      do_search_sequence(found_x, found_y, "no")
      do_search_and_copy_delivery_date(found_x, found_y)
    }
  }
}


do_search_and_copy_delivery_date(found_x, found_y)
{
  do_find_and_click_pixel(found_x, found_y, 535, 367, 1088, 526, 0x9898E1, "yes")
  do_find_and_click_image(found_x, found_y, 0, 0, 1920, 1040, "C:\Users\vnsta\AppData\Roaming\ImageRecognition\ImageOfHeader_EmailLogin.png", "yes")
  result := ocr_from_rect(508, 206, 133, 23)
  delivery_date := SubStr(result.Text, 7, 10)
  A_Clipboard := delivery_date
}

do_search_sequence(found_x, found_y, is_sequence_one)
{
  do_find_and_click_image(found_x, found_y, 0, 0, 1920, 1040, "C:\Users\vnsta\AppData\Roaming\ImageRecognition\ImageOfSearchButtonHyperlink_EmailLogin.png", "yes")
  do_find_and_click_image(found_x, found_y, 0, 0, 1920, 1040, "C:\Users\vnsta\AppData\Roaming\ImageRecognition\ImageOfSearchbarFocused_EmailLogin.png", "yes")
  SendText(A_Clipboard)

  If (is_sequence_one = "yes")
  {
    do_find_and_click_image(found_x, found_y, 0, 0, 1920, 1040, "C:\Users\vnsta\AppData\Roaming\ImageRecognition\ImageOfLeftDropbox_EmailLogin.png", "yes")
    SendInput("{Down}{Enter}") ; navigate dropbox   
  }
  do_find_and_click_image(found_x, found_y, 0, 0, 1920, 1040, "C:\Users\vnsta\AppData\Roaming\ImageRecognition\ImageOfRightDropbox_EmailLogin.png", "yes")
  SendInput("{Up}{Enter}") ; navigate dropbox
  do_find_and_click_image(found_x, found_y, 0, 0, 1920, 1040, "C:\Users\vnsta\AppData\Roaming\ImageRecognition\ImageOfSearchButton_EmailLogin.png", "yes")
}

do_find_and_click_image(found_x, found_y, x1, y1, x2, y2, image_file, search_until_found)
{
  loop
  {
    image_found := ImageSearch(&found_x, &found_y, x1, y1, x2, y2, image_file)
    If (image_found = 1)
    {
      center_img_srch_coords(image_file, &found_x, &found_y)
      Click(found_x, found_y, "Left", 1)
      Sleep 50
      break
    }
  } until (search_until_found = "no")
}

do_find_and_click_pixel(found_x, found_y, x1, y1, x2, y2, color, search_until_found)
{
  loop 
  {
    pixel_found := PixelSearch(&found_x, &found_y, x1, y1, x2, y2, color)
    If (pixel_found = 1)
    {
      Click(found_x, found_y, "Left", 1)
      Sleep 50
      break
    }
  } until (search_until_found = "no")
}

;UTILITIES---------------------------------------------------------
center_img_srch_coords(file, &coord_x, &coord_y)
{
  _gui := Gui()
  pic := _gui.Add("Picture",, file)

  width := 0
  height := 0
	pic.GetPos(,, &width, &height)
	coord_x += width // 2
	coord_y += height // 2
}
;-------------------------------------------------------------------

main()
{
  login_page_name := "HS Merseburg - Login"
  web_page_name := "HS Merseburg Webmail 1.4.23 [SVN]"
  SetTitleMatchMode(2)
  CoordMode("Pixel", "Window")
  CoordMode("Mouse", "Window")

  try_open_login_page(login_page_name)
  do_login(login_page_name)
  do_search(web_page_name)
}


; auto-execute section
main()
